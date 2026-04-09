using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace libmsstyle
{
    /// <summary>
    /// Represents color scheme info parsed from THEMES_INI.
    /// </summary>
    public class XpColorScheme
    {
        public string InternalName { get; set; } // e.g. "NormalColor", "Metallic", "HomeStead"
        public string DisplayName { get; set; }  // e.g. "Default (blue)", "Silver", "Olive Green"
        public string IniResourceName { get; set; } // e.g. "NORMALBLUE_INI"
    }

    /// <summary>
    /// Parses THEMES_INI and scheme-specific INI resources from Windows XP .msstyles files.
    /// </summary>
    public static class XpStyleIniParser
    {
        // Maps INI property names (lowercase) to (nameID, typeID) pairs.
        private static readonly Dictionary<string, (int nameID, int typeID)> s_propertyNameMap;

        // Maps INI enum value names (lowercase) to integer values, keyed by nameID.
        private static readonly Dictionary<int, Dictionary<string, int>> s_enumValueMap;

        static XpStyleIniParser()
        {
            s_propertyNameMap = BuildPropertyNameMap();
            s_enumValueMap = BuildEnumValueMap();
        }

        /// <summary>
        /// Parses UTF-16LE THEMES_INI resource to extract color scheme information.
        /// Always filters for NormalSize only.
        /// </summary>
        public static List<XpColorScheme> ParseThemesIni(byte[] data)
        {
            string text = DecodeUtf16LE(data);
            var sections = ParseIniSections(text);

            // Parse [ColorScheme.*] sections to get scheme names + display names
            var schemes = new Dictionary<string, XpColorScheme>(StringComparer.OrdinalIgnoreCase);
            foreach (var section in sections)
            {
                if (section.Key.StartsWith("ColorScheme.", StringComparison.OrdinalIgnoreCase))
                {
                    string schemeName = section.Key.Substring("ColorScheme.".Length);
                    string displayName = schemeName;
                    if (section.Value.TryGetValue("displayname", out string dn))
                        displayName = dn;

                    schemes[schemeName] = new XpColorScheme
                    {
                        InternalName = schemeName,
                        DisplayName = displayName,
                    };
                }
            }

            // Parse [File.*] sections to map scheme → INI resource name for NormalSize
            foreach (var section in sections)
            {
                if (!section.Key.StartsWith("File.", StringComparison.OrdinalIgnoreCase))
                    continue;

                var props = section.Value;
                if (!props.TryGetValue("sizes", out string sizes))
                    continue;
                if (!props.TryGetValue("colorschemes", out string colorScheme))
                    continue;

                // Only care about NormalSize
                if (!sizes.Equals("NormalSize", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Match by internal name first, then fall back to display name.
                // Some themes use DisplayName in ColorSchemes (e.g. "Lakrits" instead of "NormalColor").
                XpColorScheme matchedScheme = null;
                if (schemes.TryGetValue(colorScheme, out var schemeByName))
                {
                    matchedScheme = schemeByName;
                }
                else
                {
                    matchedScheme = schemes.Values.FirstOrDefault(s =>
                        s.DisplayName.Equals(colorScheme, StringComparison.OrdinalIgnoreCase));
                }

                if (matchedScheme != null)
                {
                    string fileSectionSuffix = section.Key.Substring("File.".Length);
                    matchedScheme.IniResourceName = fileSectionSuffix.ToUpperInvariant() + "_INI";
                }
            }

            // Return only schemes that have an INI resource mapped
            return schemes.Values.Where(s => !string.IsNullOrEmpty(s.IniResourceName)).ToList();
        }

        /// <summary>
        /// Parses a scheme-specific INI resource (e.g. NORMALBLUE_INI) into the style class hierarchy.
        /// </summary>
        public static void ParseSchemeIni(byte[] data, Dictionary<int, StyleClass> classes)
        {
            string text = DecodeUtf16LE(data);
            var sections = ParseIniSections(text);

            // Track class/part/state assignments. XP uses text names, we assign numeric IDs.
            // className (case-insensitive) → StyleClass
            var classLookup = new Dictionary<string, StyleClass>(StringComparer.OrdinalIgnoreCase);
            int nextClassId = 0;

            foreach (var section in sections)
            {
                ParseSection(section.Key, section.Value, classes, classLookup, ref nextClassId);
            }
        }

        private static void ParseSection(
            string sectionName,
            Dictionary<string, string> properties,
            Dictionary<int, StyleClass> classes,
            Dictionary<string, StyleClass> classLookup,
            ref int nextClassId)
        {
            // Parse section name: [class], [class.part], [class.part(state)], [class(state)]
            // Also: [scope::class.part(state)]
            string className;
            string partName = null;
            string stateName = null;

            // Extract state name from parentheses
            string sectionBody = sectionName;
            int parenOpen = sectionBody.IndexOf('(');
            if (parenOpen >= 0)
            {
                int parenClose = sectionBody.IndexOf(')', parenOpen);
                if (parenClose > parenOpen)
                {
                    stateName = sectionBody.Substring(parenOpen + 1, parenClose - parenOpen - 1).Trim();
                    sectionBody = sectionBody.Substring(0, parenOpen).Trim();
                }
            }

            // Split by '.' to get class and part
            int dotIndex = sectionBody.IndexOf('.');
            if (dotIndex >= 0)
            {
                className = sectionBody.Substring(0, dotIndex).Trim();
                partName = sectionBody.Substring(dotIndex + 1).Trim();
            }
            else
            {
                className = sectionBody.Trim();
            }

            if (string.IsNullOrEmpty(className))
                return;

            // Get or create StyleClass
            if (!classLookup.TryGetValue(className, out StyleClass styleClass))
            {
                styleClass = new StyleClass
                {
                    ClassId = nextClassId,
                    ClassName = className,
                };
                classLookup[className] = styleClass;
                classes[nextClassId] = styleClass;
                nextClassId++;
            }

            // Determine part ID
            int partId = 0; // 0 = common properties / class-level
            if (!string.IsNullOrEmpty(partName))
            {
                // Look for existing part with this name
                var existingPart = styleClass.Parts.Values.FirstOrDefault(
                    p => p.PartName.Equals(partName, StringComparison.OrdinalIgnoreCase));
                if (existingPart != null)
                {
                    partId = existingPart.PartId;
                }
                else
                {
                    // Assign next available part ID (starting from 1, since 0 = common)
                    partId = styleClass.Parts.Count > 0 ? styleClass.Parts.Keys.Max() + 1 : 1;
                }
            }

            // Get or create StylePart
            if (!styleClass.Parts.TryGetValue(partId, out StylePart stylePart))
            {
                stylePart = new StylePart
                {
                    PartId = partId,
                    PartName = partName ?? "Common Properties",
                };
                styleClass.Parts[partId] = stylePart;
            }

            // Determine state ID
            int stateId = 0; // 0 = default/no state
            if (!string.IsNullOrEmpty(stateName))
            {
                var existingState = stylePart.States.Values.FirstOrDefault(
                    s => s.StateName.Equals(stateName, StringComparison.OrdinalIgnoreCase));
                if (existingState != null)
                {
                    stateId = existingState.StateId;
                }
                else
                {
                    stateId = stylePart.States.Count > 0 ? stylePart.States.Keys.Max() + 1 : 1;
                }
            }

            // Get or create StyleState
            if (!stylePart.States.TryGetValue(stateId, out StyleState styleState))
            {
                styleState = new StyleState
                {
                    StateId = stateId,
                    StateName = stateName ?? "Common Properties",
                };
                stylePart.States[stateId] = styleState;
            }

            // Parse properties and add to state
            foreach (var kvp in properties)
            {
                var prop = ParseProperty(kvp.Key, kvp.Value, styleClass.ClassId, partId, stateId);
                if (prop != null)
                {
                    styleState.Properties.Add(prop);
                }
            }
        }

        private static StyleProperty ParseProperty(string name, string value, int classId, int partId, int stateId)
        {
            string nameLower = name.ToLowerInvariant();

            if (!s_propertyNameMap.TryGetValue(nameLower, out var mapping))
            {
                // Unknown property — skip
                return null;
            }

            int nameID = mapping.nameID;
            int typeID = mapping.typeID;

            var header = new PropertyHeader(nameID, typeID);
            header.classID = classId;
            header.partID = partId;
            header.stateID = stateId;

            var prop = new StyleProperty(header);

            try
            {
                switch ((IDENTIFIER)typeID)
                {
                    case IDENTIFIER.ENUM:
                        prop.SetValue(ParseEnumValue(nameID, value));
                        break;
                    case IDENTIFIER.STRING:
                        prop.SetValue(value);
                        break;
                    case IDENTIFIER.INT:
                    case IDENTIFIER.SIZE:
                        prop.SetValue(ParseInt(value));
                        break;
                    case IDENTIFIER.BOOLTYPE:
                        prop.SetValue(ParseBool(value));
                        break;
                    case IDENTIFIER.COLOR:
                        prop.SetValue(ParseColor(value));
                        break;
                    case IDENTIFIER.MARGINS:
                        prop.SetValue(ParseMargins(value));
                        break;
                    case IDENTIFIER.FILENAME:
                        // For XP, store the filename string as string value.
                        // The resource ID mapping happens at load time.
                        prop.SetValue(value.Trim());
                        break;
                    case IDENTIFIER.POSITION:
                        prop.SetValue(ParsePosition(value));
                        break;
                    case IDENTIFIER.RECTTYPE:
                        prop.SetValue(ParseMargins(value)); // same format as margins
                        break;
                    case IDENTIFIER.FONT:
                        // XP fonts are specified as "FontName, Size[, bold|italic]"
                        // Store as string for now — font handling in string table
                        prop.SetValue(value.Trim());
                        break;
                    default:
                        prop.SetValue(value.Trim());
                        break;
                }
            }
            catch
            {
                // If parsing fails, store raw string value
                prop.SetValue(value.Trim());
            }

            return prop;
        }

        #region Value Parsers

        private static int ParseEnumValue(int nameID, string value)
        {
            string trimmed = value.Trim().ToLowerInvariant();

            if (s_enumValueMap.TryGetValue(nameID, out var enumMap))
            {
                if (enumMap.TryGetValue(trimmed, out int enumVal))
                    return enumVal;
            }

            // Try parsing as integer
            if (int.TryParse(trimmed, out int intVal))
                return intVal;

            return 0;
        }

        private static int ParseInt(string value)
        {
            string trimmed = value.Trim();
            // Remove any trailing comment
            int semicolon = trimmed.IndexOf(';');
            if (semicolon >= 0)
                trimmed = trimmed.Substring(0, semicolon).Trim();

            if (int.TryParse(trimmed, out int result))
                return result;
            return 0;
        }

        private static bool ParseBool(string value)
        {
            string trimmed = value.Trim().ToLowerInvariant();
            return trimmed == "true" || trimmed == "1";
        }

        private static Color ParseColor(string value)
        {
            string trimmed = value.Trim();
            // Remove trailing comment
            int semicolon = trimmed.IndexOf(';');
            if (semicolon >= 0)
                trimmed = trimmed.Substring(0, semicolon).Trim();

            // Format: "R G B" or "R, G, B"
            var parts = trimmed.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 3)
            {
                int r = int.Parse(parts[0].Trim());
                int g = int.Parse(parts[1].Trim());
                int b = int.Parse(parts[2].Trim());
                return Color.FromArgb(r, g, b);
            }
            return Color.Black;
        }

        private static Margins ParseMargins(string value)
        {
            string trimmed = value.Trim();
            int semicolon = trimmed.IndexOf(';');
            if (semicolon >= 0)
                trimmed = trimmed.Substring(0, semicolon).Trim();

            var parts = trimmed.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 4)
            {
                return new Margins(
                    int.Parse(parts[0].Trim()),
                    int.Parse(parts[1].Trim()),
                    int.Parse(parts[2].Trim()),
                    int.Parse(parts[3].Trim()));
            }
            return new Margins(0, 0, 0, 0);
        }

        private static System.Drawing.Size ParsePosition(string value)
        {
            string trimmed = value.Trim();
            int semicolon = trimmed.IndexOf(';');
            if (semicolon >= 0)
                trimmed = trimmed.Substring(0, semicolon).Trim();

            var parts = trimmed.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2)
            {
                return new System.Drawing.Size(
                    int.Parse(parts[0].Trim()),
                    int.Parse(parts[1].Trim()));
            }
            return new System.Drawing.Size(0, 0);
        }

        #endregion

        #region INI Helper Methods

        private static string DecodeUtf16LE(byte[] data)
        {
            if (data == null || data.Length < 2)
                return string.Empty;

            // Check for BOM
            int start = 0;
            if (data[0] == 0xFF && data[1] == 0xFE)
                start = 2;

            return Encoding.Unicode.GetString(data, start, data.Length - start);
        }

        /// <summary>
        /// Parses INI text into sections. Each section maps to a dictionary of key-value pairs.
        /// Keys are stored lowercase for case-insensitive matching.
        /// </summary>
        private static Dictionary<string, Dictionary<string, string>> ParseIniSections(string text)
        {
            var sections = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            string currentSection = null;
            Dictionary<string, string> currentProps = null;

            using (var reader = new StringReader(text))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    line = line.Trim();

                    // Skip empty lines and comments
                    if (string.IsNullOrEmpty(line) || line.StartsWith(";"))
                        continue;

                    // Section header
                    if (line.StartsWith("[") && line.Contains("]"))
                    {
                        int end = line.IndexOf(']');
                        currentSection = line.Substring(1, end - 1).Trim();
                        if (!sections.TryGetValue(currentSection, out currentProps))
                        {
                            currentProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            sections[currentSection] = currentProps;
                        }
                        continue;
                    }

                    // Key = Value
                    if (currentSection != null)
                    {
                        int eq = line.IndexOf('=');
                        if (eq > 0)
                        {
                            string key = line.Substring(0, eq).Trim();
                            string val = line.Substring(eq + 1).Trim();
                            currentProps[key] = val;
                        }
                    }
                }
            }

            return sections;
        }

        #endregion

        #region Image Resource Name Mapping

        /// <summary>
        /// Converts an INI image file reference (e.g. "Blue\CloseButton.bmp") to a PE BITMAP resource name.
        /// Convention from luna.rc: path separators → '_', remove extension, uppercase, prefix with scheme prefix.
        /// E.g. "Blue\CloseButton.bmp" → "BLUE_CLOSEBUTTON_BMP"
        /// </summary>
        public static string ImageFileToResourceName(string imageFile)
        {
            if (string.IsNullOrEmpty(imageFile))
                return string.Empty;

            // Remove quotes if present
            imageFile = imageFile.Trim().Trim('"');

            // Replace path separators with underscore, remove extension
            string name = imageFile
                .Replace('\\', '_')
                .Replace('/', '_');

            // Replace the '.' before extension with '_'
            int lastDot = name.LastIndexOf('.');
            if (lastDot >= 0)
            {
                name = name.Substring(0, lastDot) + "_" + name.Substring(lastDot + 1);
            }

            return name.ToUpperInvariant();
        }

        #endregion

        #region Property Name → IDENTIFIER Mapping

        private static Dictionary<string, (int nameID, int typeID)> BuildPropertyNameMap()
        {
            // Map XP INI property names (lowercase) to (nameID, typeID).
            // Based on VisualStyleDefinitions IDENTIFIER enum values and VisualStylePropertyMap.
            var map = new Dictionary<string, (int, int)>(StringComparer.OrdinalIgnoreCase)
            {
                // Enum properties
                ["bgtype"] = ((int)IDENTIFIER.BGTYPE, (int)IDENTIFIER.ENUM),
                ["bordertype"] = ((int)IDENTIFIER.BORDERTYPE, (int)IDENTIFIER.ENUM),
                ["filltype"] = ((int)IDENTIFIER.FILLTYPE, (int)IDENTIFIER.ENUM),
                ["sizingtype"] = ((int)IDENTIFIER.SIZINGTYPE, (int)IDENTIFIER.ENUM),
                ["halign"] = ((int)IDENTIFIER.HALIGN, (int)IDENTIFIER.ENUM),
                ["contentalignment"] = ((int)IDENTIFIER.CONTENTALIGNMENT, (int)IDENTIFIER.ENUM),
                ["valign"] = ((int)IDENTIFIER.VALIGN, (int)IDENTIFIER.ENUM),
                ["offsettype"] = ((int)IDENTIFIER.OFFSETTYPE, (int)IDENTIFIER.ENUM),
                ["iconeffect"] = ((int)IDENTIFIER.ICONEFFECT, (int)IDENTIFIER.ENUM),
                ["textshadowtype"] = ((int)IDENTIFIER.TEXTSHADOWTYPE, (int)IDENTIFIER.ENUM),
                ["imagelayout"] = ((int)IDENTIFIER.IMAGELAYOUT, (int)IDENTIFIER.ENUM),
                ["glyphtype"] = ((int)IDENTIFIER.GLYPHTYPE, (int)IDENTIFIER.ENUM),
                ["imageselecttype"] = ((int)IDENTIFIER.IMAGESELECTTYPE, (int)IDENTIFIER.ENUM),
                ["glyphfontsizingtype"] = ((int)IDENTIFIER.GLYPHFONTSIZINGTYPE, (int)IDENTIFIER.ENUM),
                ["truesizescalingtype"] = ((int)IDENTIFIER.TRUESIZESCALINGTYPE, (int)IDENTIFIER.ENUM),

                // Bool properties
                ["transparent"] = ((int)IDENTIFIER.TRANSPARENT_, (int)IDENTIFIER.BOOLTYPE),
                ["autosize"] = ((int)IDENTIFIER.AUTOSIZE, (int)IDENTIFIER.BOOLTYPE),
                ["borderonly"] = ((int)IDENTIFIER.BORDERONLY, (int)IDENTIFIER.BOOLTYPE),
                ["composited"] = ((int)IDENTIFIER.COMPOSITED, (int)IDENTIFIER.BOOLTYPE),
                ["bgfill"] = ((int)IDENTIFIER.BGFILL, (int)IDENTIFIER.BOOLTYPE),
                ["glyphtransparent"] = ((int)IDENTIFIER.GLYPHTRANSPARENT, (int)IDENTIFIER.BOOLTYPE),
                ["glyphonly"] = ((int)IDENTIFIER.GLYPHONLY, (int)IDENTIFIER.BOOLTYPE),
                ["alwaysshowsizingbar"] = ((int)IDENTIFIER.ALWAYSSHOWSIZINGBAR, (int)IDENTIFIER.BOOLTYPE),
                ["mirrorimage"] = ((int)IDENTIFIER.MIRRORIMAGE, (int)IDENTIFIER.BOOLTYPE),
                ["uniformsizing"] = ((int)IDENTIFIER.UNIFORMSIZING, (int)IDENTIFIER.BOOLTYPE),
                ["integralsizing"] = ((int)IDENTIFIER.INTEGRALSIZING, (int)IDENTIFIER.BOOLTYPE),
                ["sourcegrow"] = ((int)IDENTIFIER.SOURCEGROW, (int)IDENTIFIER.BOOLTYPE),
                ["sourceshrink"] = ((int)IDENTIFIER.SOURCESHRINK, (int)IDENTIFIER.BOOLTYPE),
                ["userpicture"] = ((int)IDENTIFIER.USERPICTURE, (int)IDENTIFIER.BOOLTYPE),
                ["flatmenus"] = ((int)IDENTIFIER.FLATMENUS, (int)IDENTIFIER.BOOLTYPE),
                ["localizedmirrorimage"] = ((int)IDENTIFIER.LOCALIZEDMIRRORIMAGE, (int)IDENTIFIER.BOOLTYPE),

                // Int properties
                ["imagecount"] = ((int)IDENTIFIER.IMAGECOUNT, (int)IDENTIFIER.INT),
                ["alphalevel"] = ((int)IDENTIFIER.ALPHALEVEL, (int)IDENTIFIER.INT),
                ["bordersize"] = ((int)IDENTIFIER.BORDERSIZE, (int)IDENTIFIER.INT),
                ["roundcornerwidth"] = ((int)IDENTIFIER.ROUNDCORNERWIDTH, (int)IDENTIFIER.INT),
                ["roundcornerheight"] = ((int)IDENTIFIER.ROUNDCORNERHEIGHT, (int)IDENTIFIER.INT),
                ["gradientratio1"] = ((int)IDENTIFIER.GRADIENTRATIO1, (int)IDENTIFIER.INT),
                ["gradientratio2"] = ((int)IDENTIFIER.GRADIENTRATIO2, (int)IDENTIFIER.INT),
                ["gradientratio3"] = ((int)IDENTIFIER.GRADIENTRATIO3, (int)IDENTIFIER.INT),
                ["gradientratio4"] = ((int)IDENTIFIER.GRADIENTRATIO4, (int)IDENTIFIER.INT),
                ["gradientratio5"] = ((int)IDENTIFIER.GRADIENTRATIO5, (int)IDENTIFIER.INT),
                ["progresschunksize"] = ((int)IDENTIFIER.PROGRESSCHUNKSIZE, (int)IDENTIFIER.INT),
                ["progressspacesize"] = ((int)IDENTIFIER.PROGRESSSPACESIZE, (int)IDENTIFIER.INT),
                ["saturation"] = ((int)IDENTIFIER.SATURATION, (int)IDENTIFIER.INT),
                ["textbordersize"] = ((int)IDENTIFIER.TEXTBORDERSIZE, (int)IDENTIFIER.INT),
                ["alphathreshold"] = ((int)IDENTIFIER.ALPHATHRESHOLD, (int)IDENTIFIER.INT),
                ["width"] = ((int)IDENTIFIER.WIDTH, (int)IDENTIFIER.INT),
                ["height"] = ((int)IDENTIFIER.HEIGHT, (int)IDENTIFIER.INT),
                ["glyphindex"] = ((int)IDENTIFIER.GLYPHINDEX, (int)IDENTIFIER.INT),
                ["truesizestretchmark"] = ((int)IDENTIFIER.TRUESIZESTRETCHMARK, (int)IDENTIFIER.INT),
                ["mindpi1"] = ((int)IDENTIFIER.MINDPI1, (int)IDENTIFIER.INT),
                ["mindpi2"] = ((int)IDENTIFIER.MINDPI2, (int)IDENTIFIER.INT),
                ["mindpi3"] = ((int)IDENTIFIER.MINDPI3, (int)IDENTIFIER.INT),
                ["mindpi4"] = ((int)IDENTIFIER.MINDPI4, (int)IDENTIFIER.INT),
                ["mindpi5"] = ((int)IDENTIFIER.MINDPI5, (int)IDENTIFIER.INT),
                ["textglowsize"] = ((int)IDENTIFIER.TEXTGLOWSIZE, (int)IDENTIFIER.INT),
                ["framespersecond"] = ((int)IDENTIFIER.FRAMESPERSECOND, (int)IDENTIFIER.INT),
                ["pixelsperframe"] = ((int)IDENTIFIER.PIXELSPERFRAME, (int)IDENTIFIER.INT),
                ["animationdelay"] = ((int)IDENTIFIER.ANIMATIONDELAY, (int)IDENTIFIER.INT),
                ["glowintensity"] = ((int)IDENTIFIER.GLOWINTENSITY, (int)IDENTIFIER.INT),
                ["opacity"] = ((int)IDENTIFIER.OPACITY, (int)IDENTIFIER.INT),
                ["mincolordepth"] = ((int)IDENTIFIER.MINCOLORDEPTH, (int)IDENTIFIER.INT),

                // Size properties (uses SIZE type = 207, same as INT in binary)
                ["sizingborderwidth"] = ((int)IDENTIFIER.SIZINGBORDERWIDTH, (int)IDENTIFIER.SIZE),
                ["scrollbarwidth"] = ((int)IDENTIFIER.SCROLLBARWIDTH, (int)IDENTIFIER.SIZE),
                ["scrollbarheight"] = ((int)IDENTIFIER.SCROLLBARHEIGHT, (int)IDENTIFIER.SIZE),
                ["captionbarwidth"] = ((int)IDENTIFIER.CAPTIONBARWIDTH, (int)IDENTIFIER.SIZE),
                ["captionbarheight"] = ((int)IDENTIFIER.CAPTIONBARHEIGHT, (int)IDENTIFIER.SIZE),
                ["smcaptionbarwidth"] = ((int)IDENTIFIER.SMCAPTIONBARWIDTH, (int)IDENTIFIER.SIZE),
                ["smcaptionbarheight"] = ((int)IDENTIFIER.SMCAPTIONBARHEIGHT, (int)IDENTIFIER.SIZE),
                ["menubarwidth"] = ((int)IDENTIFIER.MENUBARWIDTH, (int)IDENTIFIER.SIZE),
                ["menubarheight"] = ((int)IDENTIFIER.MENUBARHEIGHT, (int)IDENTIFIER.SIZE),

                // Color properties
                ["bordercolor"] = ((int)IDENTIFIER.BORDERCOLOR, (int)IDENTIFIER.COLOR),
                ["fillcolor"] = ((int)IDENTIFIER.FILLCOLOR, (int)IDENTIFIER.COLOR),
                ["textcolor"] = ((int)IDENTIFIER.TEXTCOLOR, (int)IDENTIFIER.COLOR),
                ["edgelightcolor"] = ((int)IDENTIFIER.EDGELIGHTCOLOR, (int)IDENTIFIER.COLOR),
                ["edgehighlightcolor"] = ((int)IDENTIFIER.EDGEHIGHLIGHTCOLOR, (int)IDENTIFIER.COLOR),
                ["edgeshadowcolor"] = ((int)IDENTIFIER.EDGESHADOWCOLOR, (int)IDENTIFIER.COLOR),
                ["edgedkshadowcolor"] = ((int)IDENTIFIER.EDGEDKSHADOWCOLOR, (int)IDENTIFIER.COLOR),
                ["edgefillcolor"] = ((int)IDENTIFIER.EDGEFILLCOLOR, (int)IDENTIFIER.COLOR),
                ["transparentcolor"] = ((int)IDENTIFIER.TRANSPARENTCOLOR, (int)IDENTIFIER.COLOR),
                ["gradientcolor1"] = ((int)IDENTIFIER.GRADIENTCOLOR1, (int)IDENTIFIER.COLOR),
                ["gradientcolor2"] = ((int)IDENTIFIER.GRADIENTCOLOR2, (int)IDENTIFIER.COLOR),
                ["gradientcolor3"] = ((int)IDENTIFIER.GRADIENTCOLOR3, (int)IDENTIFIER.COLOR),
                ["gradientcolor4"] = ((int)IDENTIFIER.GRADIENTCOLOR4, (int)IDENTIFIER.COLOR),
                ["gradientcolor5"] = ((int)IDENTIFIER.GRADIENTCOLOR5, (int)IDENTIFIER.COLOR),
                ["shadowcolor"] = ((int)IDENTIFIER.SHADOWCOLOR, (int)IDENTIFIER.COLOR),
                ["glowcolor"] = ((int)IDENTIFIER.GLOWCOLOR, (int)IDENTIFIER.COLOR),
                ["textbordercolor"] = ((int)IDENTIFIER.TEXTBORDERCOLOR, (int)IDENTIFIER.COLOR),
                ["textshadowcolor"] = ((int)IDENTIFIER.TEXTSHADOWCOLOR, (int)IDENTIFIER.COLOR),
                ["glyphtextcolor"] = ((int)IDENTIFIER.GLYPHTEXTCOLOR, (int)IDENTIFIER.COLOR),
                ["glyphtransparentcolor"] = ((int)IDENTIFIER.GLYPHTRANSPARENTCOLOR, (int)IDENTIFIER.COLOR),
                ["fillcolorhint"] = ((int)IDENTIFIER.FILLCOLORHINT, (int)IDENTIFIER.COLOR),
                ["bordercolorhint"] = ((int)IDENTIFIER.BORDERCOLORHINT, (int)IDENTIFIER.COLOR),
                ["accentcolorhint"] = ((int)IDENTIFIER.ACCENTCOLORHINT, (int)IDENTIFIER.COLOR),
                ["textcolorhint"] = ((int)IDENTIFIER.TEXTCOLORHINT, (int)IDENTIFIER.COLOR),
                ["heading1textcolor"] = ((int)IDENTIFIER.HEADING1TEXTCOLOR, (int)IDENTIFIER.COLOR),
                ["heading2textcolor"] = ((int)IDENTIFIER.HEADING2TEXTCOLOR, (int)IDENTIFIER.COLOR),
                ["bodytextcolor"] = ((int)IDENTIFIER.BODYTEXTCOLOR, (int)IDENTIFIER.COLOR),
                ["blendcolor"] = ((int)IDENTIFIER.BLENDCOLOR, (int)IDENTIFIER.COLOR),

                // System metric colors (globals section)
                ["scrollbar"] = ((int)IDENTIFIER.SCROLLBAR, (int)IDENTIFIER.COLOR),
                ["background"] = ((int)IDENTIFIER.BACKGROUND, (int)IDENTIFIER.COLOR),
                ["activecaption"] = ((int)IDENTIFIER.ACTIVECAPTION, (int)IDENTIFIER.COLOR),
                ["inactivecaption"] = ((int)IDENTIFIER.INACTIVECAPTION, (int)IDENTIFIER.COLOR),
                ["menu"] = ((int)IDENTIFIER.MENU, (int)IDENTIFIER.COLOR),
                ["window"] = ((int)IDENTIFIER.WINDOW, (int)IDENTIFIER.COLOR),
                ["windowframe"] = ((int)IDENTIFIER.WINDOWFRAME, (int)IDENTIFIER.COLOR),
                ["menutext"] = ((int)IDENTIFIER.MENUTEXT, (int)IDENTIFIER.COLOR),
                ["windowtext"] = ((int)IDENTIFIER.WINDOWTEXT, (int)IDENTIFIER.COLOR),
                ["captiontext"] = ((int)IDENTIFIER.CAPTIONTEXT, (int)IDENTIFIER.COLOR),
                ["activeborder"] = ((int)IDENTIFIER.ACTIVEBORDER, (int)IDENTIFIER.COLOR),
                ["inactiveborder"] = ((int)IDENTIFIER.INACTIVEBORDER, (int)IDENTIFIER.COLOR),
                ["appworkspace"] = ((int)IDENTIFIER.APPWORKSPACE, (int)IDENTIFIER.COLOR),
                ["highlight"] = ((int)IDENTIFIER.HIGHLIGHT, (int)IDENTIFIER.COLOR),
                ["highlighttext"] = ((int)IDENTIFIER.HIGHLIGHTTEXT, (int)IDENTIFIER.COLOR),
                ["btnface"] = ((int)IDENTIFIER.BTNFACE, (int)IDENTIFIER.COLOR),
                ["btnshadow"] = ((int)IDENTIFIER.BTNSHADOW, (int)IDENTIFIER.COLOR),
                ["graytext"] = ((int)IDENTIFIER.GRAYTEXT, (int)IDENTIFIER.COLOR),
                ["btntext"] = ((int)IDENTIFIER.BTNTEXT, (int)IDENTIFIER.COLOR),
                ["inactivecaptiontext"] = ((int)IDENTIFIER.INACTIVECAPTIONTEXT, (int)IDENTIFIER.COLOR),
                ["btnhighlight"] = ((int)IDENTIFIER.BTNHIGHLIGHT, (int)IDENTIFIER.COLOR),
                ["dkshadow3d"] = ((int)IDENTIFIER.DKSHADOW3D, (int)IDENTIFIER.COLOR),
                ["light3d"] = ((int)IDENTIFIER.LIGHT3D, (int)IDENTIFIER.COLOR),
                ["infotext"] = ((int)IDENTIFIER.INFOTEXT, (int)IDENTIFIER.COLOR),
                ["infobk"] = ((int)IDENTIFIER.INFOBK, (int)IDENTIFIER.COLOR),
                ["buttonalternateface"] = ((int)IDENTIFIER.BUTTONALTERNATEFACE, (int)IDENTIFIER.COLOR),
                ["hottracking"] = ((int)IDENTIFIER.HOTTRACKING, (int)IDENTIFIER.COLOR),
                ["gradientactivecaption"] = ((int)IDENTIFIER.GRADIENTACTIVECAPTION, (int)IDENTIFIER.COLOR),
                ["gradientinactivecaption"] = ((int)IDENTIFIER.GRADIENTINACTIVECAPTION, (int)IDENTIFIER.COLOR),
                ["menuhilight"] = ((int)IDENTIFIER.MENUHILIGHT, (int)IDENTIFIER.COLOR),
                ["menubar"] = ((int)IDENTIFIER.MENUBAR, (int)IDENTIFIER.COLOR),

                // Filename / image properties
                ["imagefile"] = ((int)IDENTIFIER.IMAGEFILE, (int)IDENTIFIER.FILENAME),
                ["imagefile1"] = ((int)IDENTIFIER.IMAGEFILE1, (int)IDENTIFIER.FILENAME),
                ["imagefile2"] = ((int)IDENTIFIER.IMAGEFILE2, (int)IDENTIFIER.FILENAME),
                ["imagefile3"] = ((int)IDENTIFIER.IMAGEFILE3, (int)IDENTIFIER.FILENAME),
                ["imagefile4"] = ((int)IDENTIFIER.IMAGEFILE4, (int)IDENTIFIER.FILENAME),
                ["imagefile5"] = ((int)IDENTIFIER.IMAGEFILE5, (int)IDENTIFIER.FILENAME),
                ["glyphimagefile"] = ((int)IDENTIFIER.GLYPHIMAGEFILE, (int)IDENTIFIER.FILENAME),
                ["stockimagefile"] = ((int)IDENTIFIER.IMAGEFILE, (int)IDENTIFIER.FILENAME), // alias

                // Position properties
                ["offset"] = ((int)IDENTIFIER.OFFSET, (int)IDENTIFIER.POSITION),
                ["textshadowoffset"] = ((int)IDENTIFIER.TEXTSHADOWOFFSET, (int)IDENTIFIER.POSITION),
                ["minsize"] = ((int)IDENTIFIER.MINSIZE, (int)IDENTIFIER.POSITION),
                ["minsize1"] = ((int)IDENTIFIER.MINSIZE1, (int)IDENTIFIER.POSITION),
                ["minsize2"] = ((int)IDENTIFIER.MINSIZE2, (int)IDENTIFIER.POSITION),
                ["minsize3"] = ((int)IDENTIFIER.MINSIZE3, (int)IDENTIFIER.POSITION),
                ["minsize4"] = ((int)IDENTIFIER.MINSIZE4, (int)IDENTIFIER.POSITION),
                ["minsize5"] = ((int)IDENTIFIER.MINSIZE5, (int)IDENTIFIER.POSITION),
                ["normalsize"] = ((int)IDENTIFIER.NORMALSIZE, (int)IDENTIFIER.POSITION),

                // Margins properties
                ["sizingmargins"] = ((int)IDENTIFIER.SIZINGMARGINS, (int)IDENTIFIER.MARGINS),
                ["contentmargins"] = ((int)IDENTIFIER.CONTENTMARGINS, (int)IDENTIFIER.MARGINS),
                ["captionmargins"] = ((int)IDENTIFIER.CAPTIONMARGINS, (int)IDENTIFIER.MARGINS),

                // Rect properties
                ["defaultpanesize"] = ((int)IDENTIFIER.DEFAULTPANESIZE, (int)IDENTIFIER.RECTTYPE),
                ["customsplitrect"] = ((int)IDENTIFIER.CUSTOMSPLITRECT, (int)IDENTIFIER.RECTTYPE),
                ["animationbuttonrect"] = ((int)IDENTIFIER.ANIMATIONBUTTONRECT, (int)IDENTIFIER.RECTTYPE),

                // Font properties
                ["font"] = ((int)IDENTIFIER.FONT, (int)IDENTIFIER.FONT),
                ["captionfont"] = ((int)IDENTIFIER.CAPTIONFONT, (int)IDENTIFIER.FONT),
                ["smallcaptionfont"] = ((int)IDENTIFIER.SMALLCAPTIONFONT, (int)IDENTIFIER.FONT),
                ["menufont"] = ((int)IDENTIFIER.MENUFONT, (int)IDENTIFIER.FONT),
                ["statusfont"] = ((int)IDENTIFIER.STATUSFONT, (int)IDENTIFIER.FONT),
                ["msgboxfont"] = ((int)IDENTIFIER.MSGBOXFONT, (int)IDENTIFIER.FONT),
                ["icontitlefont"] = ((int)IDENTIFIER.ICONTITLEFONT, (int)IDENTIFIER.FONT),
                ["heading1font"] = ((int)IDENTIFIER.HEADING1FONT, (int)IDENTIFIER.FONT),
                ["heading2font"] = ((int)IDENTIFIER.HEADING2FONT, (int)IDENTIFIER.FONT),
                ["bodyfont"] = ((int)IDENTIFIER.BODYFONT, (int)IDENTIFIER.FONT),
                ["glyphfont"] = ((int)IDENTIFIER.GLYPHFONT, (int)IDENTIFIER.FONT),

                // String properties
                ["text"] = ((int)IDENTIFIER.TEXT, (int)IDENTIFIER.STRING),
                ["classicvalue"] = ((int)IDENTIFIER.CLASSICVALUE, (int)IDENTIFIER.STRING),
                ["cssname"] = ((int)IDENTIFIER.CSSNAME, (int)IDENTIFIER.STRING),
                ["xmlname"] = ((int)IDENTIFIER.XMLNAME, (int)IDENTIFIER.STRING),
                ["lastupdated"] = ((int)IDENTIFIER.LASTUPDATED, (int)IDENTIFIER.STRING),
                ["alias"] = ((int)IDENTIFIER.ALIAS, (int)IDENTIFIER.STRING),

                // Int properties (animation duration)
                ["animationduration"] = ((int)IDENTIFIER.ANIMATIONDURATION, (int)IDENTIFIER.INT),
            };

            return map;
        }

        private static Dictionary<int, Dictionary<string, int>> BuildEnumValueMap()
        {
            var map = new Dictionary<int, Dictionary<string, int>>();

            // BGTYPE
            map[(int)IDENTIFIER.BGTYPE] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["imagefile"] = 0,
                ["borderfill"] = 1,
                ["none"] = 2,
            };

            // IMAGELAYOUT
            map[(int)IDENTIFIER.IMAGELAYOUT] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["vertical"] = 0,
                ["horizontal"] = 1,
            };

            // BORDERTYPE
            map[(int)IDENTIFIER.BORDERTYPE] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["rect"] = 0,
                ["roundrect"] = 1,
                ["ellipse"] = 2,
            };

            // FILLTYPE
            map[(int)IDENTIFIER.FILLTYPE] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["solid"] = 0,
                ["vertgradient"] = 1,
                ["horizontalgradient"] = 2,
                ["horzgradient"] = 2,
                ["radialgradient"] = 3,
                ["tileimage"] = 4,
            };

            // SIZINGTYPE
            map[(int)IDENTIFIER.SIZINGTYPE] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["truesize"] = 0,
                ["stretch"] = 1,
                ["tile"] = 2,
            };

            // HALIGN
            map[(int)IDENTIFIER.HALIGN] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["left"] = 0,
                ["center"] = 1,
                ["right"] = 2,
            };

            // CONTENTALIGNMENT
            map[(int)IDENTIFIER.CONTENTALIGNMENT] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["left"] = 0,
                ["center"] = 1,
                ["right"] = 2,
            };

            // VALIGN
            map[(int)IDENTIFIER.VALIGN] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["top"] = 0,
                ["center"] = 1,
                ["bottom"] = 2,
            };

            // OFFSETTYPE
            map[(int)IDENTIFIER.OFFSETTYPE] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["topleft"] = 0,
                ["topright"] = 1,
                ["topmiddle"] = 2,
                ["bottomleft"] = 3,
                ["bottomright"] = 4,
                ["bottommiddle"] = 5,
                ["middleright"] = 6,
                ["leftofcaption"] = 7,
                ["rightofcaption"] = 8,
                ["leftoflastbutton"] = 9,
                ["rightoflastbutton"] = 10,
                ["abovelastbutton"] = 11,
                ["belowlastbutton"] = 12,
            };

            // ICONEFFECT
            map[(int)IDENTIFIER.ICONEFFECT] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["none"] = 0,
                ["glow"] = 1,
                ["shadow"] = 2,
                ["pulse"] = 3,
                ["alpha"] = 4,
            };

            // TEXTSHADOWTYPE
            map[(int)IDENTIFIER.TEXTSHADOWTYPE] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["none"] = 0,
                ["single"] = 1,
                ["continuous"] = 2,
            };

            // GLYPHTYPE
            map[(int)IDENTIFIER.GLYPHTYPE] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["none"] = 0,
                ["imageglyph"] = 1,
                ["fontglyph"] = 2,
            };

            // IMAGESELECTTYPE
            map[(int)IDENTIFIER.IMAGESELECTTYPE] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["none"] = 0,
                ["size"] = 1,
                ["dpi"] = 2,
            };

            // GLYPHFONTSIZINGTYPE
            map[(int)IDENTIFIER.GLYPHFONTSIZINGTYPE] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["none"] = 0,
                ["size"] = 1,
                ["dpi"] = 2,
            };

            // TRUESIZESCALINGTYPE
            map[(int)IDENTIFIER.TRUESIZESCALINGTYPE] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["none"] = 0,
                ["size"] = 1,
                ["dpi"] = 2,
            };

            return map;
        }

        #endregion
    }
}
