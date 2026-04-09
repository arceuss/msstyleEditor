using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace libmsstyle
{
    public class VisualStyle : IDisposable
    {
        private IntPtr m_moduleHandle;

        private string m_stylePath;
        public string Path
        {
            get { return m_stylePath; }
        }

        private Platform m_platform;
        public Platform Platform
        {
            get { return m_platform; }
        }

        private Dictionary<int, StyleClass> m_classes;
        public Dictionary<int, StyleClass> Classes
        {
            get { return m_classes; }
        }

        private List<TimingFunction> m_timingFunctions;
        public List<TimingFunction> TimingFunctions
        {
            get { return m_timingFunctions; }
        }

        private List<Animation> m_animations;
        public List<Animation> Animations
        {
            get { return m_animations; }
        }

        private int m_numProps;
        public int NumProperties
        {
            get { return m_numProps; }
        }

        private Dictionary<int, string> m_stringTable;
        public Dictionary<int, string> PreferredStringTable
        {
            get { return m_stringTable; }
        }

        private Dictionary<int, Dictionary<int, string>> m_stringTables = new Dictionary<int, Dictionary<int, string>>();

        public Dictionary<int, Dictionary<int, string>> StringTables
        {
            get { return m_stringTables; }
        }

        private Dictionary<StyleResource, string> m_resourceUpdates;

        private ushort m_resourceLanguage;

        private bool m_classmapDirty = false;
        public void MarkClassmapDirty() { m_classmapDirty = true; }

        private byte[] m_originalCmap;
        private int m_cmapTotalEntries;
        private int m_originalClassCount;

        private byte[] m_originalBcmap;

        // XP-specific: maps image file paths (from INI) to BITMAP resource names
        private Dictionary<string, string> m_xpImageMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // XP-specific: the selected color scheme
        private XpColorScheme m_xpColorScheme;

        // XP-specific: all available color schemes
        private List<XpColorScheme> m_xpColorSchemes;
        public List<XpColorScheme> XpColorSchemes
        {
            get { return m_xpColorSchemes; }
        }

        public VisualStyle()
        {
            m_stylePath = null;
            m_classes = new Dictionary<int, StyleClass>();
            m_stringTable = new Dictionary<int, string>();
            m_resourceUpdates = new Dictionary<StyleResource, string>();
            m_timingFunctions = new List<TimingFunction>();
            m_animations = new List<Animation>();
            m_numProps = 0;
            m_resourceLanguage = 0;
        }

        ~VisualStyle()
        {
            Dispose(false);
        }

        private bool m_disposed = false;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!m_disposed)
            {
                if (disposing)
                {
                    // managed resources
                }

                Win32Api.FreeLibrary(m_moduleHandle);
                m_moduleHandle = IntPtr.Zero;
                m_disposed = true;
            }
        }

        public void Save(string file, bool makeStandalone = false)
        {
            if (file != m_stylePath)
            {
                File.Copy(m_stylePath, file, true);
            }
            else throw new ArgumentException("Cannot overwrite the original file!");

            var updateHandle = Win32Api.BeginUpdateResource(file, false);
            if (updateHandle == IntPtr.Zero)
            {
                File.Delete(file);
                throw new IOException("Could not open the file for writing! (BeginUpdateResource)");
            }

            var moduleHandle = Win32Api.LoadLibraryEx(file, IntPtr.Zero, Win32Api.LoadLibraryFlags.LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE);
            if (moduleHandle == IntPtr.Zero)
            {
                Win32Api.EndUpdateResource(updateHandle, true);
                File.Delete(file);
                throw new IOException("Could not open the file for writing! (LoadLibraryEx)");
            }

            if (!SaveResources(moduleHandle, updateHandle))
            {
                Win32Api.FreeLibrary(moduleHandle);
                Win32Api.EndUpdateResource(updateHandle, true);
                File.Delete(file);
                throw new IOException("Could not save resources!");
            }

            if (!SaveAMap(moduleHandle, updateHandle))
            {
                Win32Api.FreeLibrary(moduleHandle);
                Win32Api.EndUpdateResource(updateHandle, true);
                File.Delete(file);
                throw new IOException("Could not save resources!");
            }

            if (m_classmapDirty)
            {
                if (!SaveClassmap(moduleHandle, updateHandle))
                {
                    Win32Api.FreeLibrary(moduleHandle);
                    Win32Api.EndUpdateResource(updateHandle, true);
                    File.Delete(file);
                    throw new IOException("Could not save class map!");
                }
            }

            if (!SaveProperties(moduleHandle, updateHandle))
            {
                Win32Api.FreeLibrary(moduleHandle);
                Win32Api.EndUpdateResource(updateHandle, true);
                File.Delete(file);
                throw new IOException("Could not save resources!");
            }

            try
            {
                SaveStringTable(moduleHandle, updateHandle, makeStandalone);
            }
            catch (Exception)
            {
                Win32Api.FreeLibrary(moduleHandle);
                Win32Api.EndUpdateResource(updateHandle, true);
                File.Delete(file);
                throw;
            }

            // close the module before calling EndUpdate(). If not
            // updating fails because the file is in use.
            Win32Api.FreeLibrary(moduleHandle);
            if (!Win32Api.EndUpdateResource(updateHandle, false))
            {
                File.Delete(file);
                throw new IOException("Could not write the changes to the file!");
            }

            // signature
            byte[] signature = new byte[]
            {
                0x72, 0x68, 0xC5, 0x5F, 0x1F, 0x00, 0xA4, 0x9B, 0xFF, 0x90, 0xB2, 0x94, 0x7F, 0x76, 0x38, 0x38,
                0x48, 0xD3, 0x9D, 0x80, 0xB2, 0x69, 0x0C, 0x52, 0xBC, 0x82, 0xDA, 0x1D, 0xC8, 0x54, 0xD4, 0xE3,
                0xD4, 0xC7, 0xA6, 0x79, 0xD1, 0x06, 0xBD, 0x44, 0xDD, 0x99, 0x57, 0x9C, 0x3E, 0xDD, 0xAD, 0xA6,
                0x58, 0x16, 0x49, 0xC7, 0x55, 0x93, 0x0E, 0xD1, 0x89, 0xB0, 0x71, 0x63, 0x2E, 0xE9, 0xDF, 0x02,
                0x26, 0x88, 0xEF, 0x56, 0x25, 0x5A, 0xA4, 0x04, 0xD5, 0xAB, 0x71, 0x31, 0xC2, 0x48, 0x29, 0xC4,
                0x13, 0xD0, 0x5B, 0x81, 0x3D, 0xCC, 0x27, 0x0A, 0xD6, 0xEE, 0x5C, 0x9E, 0x99, 0xE9, 0x53, 0x6D,
                0x87, 0x72, 0x41, 0x44, 0xAF, 0x61, 0xA0, 0x87, 0xE2, 0x3C, 0xE0, 0x62, 0x98, 0x26, 0xBF, 0xE7,
                0x80, 0xFF, 0x23, 0xCA, 0xF7, 0xC6, 0x34, 0x6C, 0x9A, 0xA8, 0xA1, 0xA6, 0xEE, 0xA4, 0xB6, 0xEE
            };
            Signature.AppendSignature(file, signature);

        }


        private bool SaveResources(IntPtr moduleHandle, IntPtr updateHandle)
        {
            foreach (var res in m_resourceUpdates)
            {
                byte[] data = null;
                try
                {
                    data = File.ReadAllBytes(res.Value);
                }
                catch (Exception)
                {
                    return false;
                }

                string type = "";
                switch (res.Key.Type)
                {
                    case StyleResourceType.None:
                        continue;
                    case StyleResourceType.Image:
                        type = "IMAGE"; break;
                    case StyleResourceType.Atlas:
                        type = "STREAM"; break;
                }

                ushort lid = ResourceAccess.GetFirstLanguageId(moduleHandle, type, (uint)res.Key.ResourceId);
                if (lid == 0xFFFF)
                {
                    lid = m_resourceLanguage;
                }
                if (!Win32Api.UpdateResource(updateHandle, type, (uint)res.Key.ResourceId, lid, data, (uint)data.Length))
                {
                    return false;
                }
            }

            return true;
        }


        private bool SaveProperties(IntPtr moduleHandle, IntPtr updateHandle)
        {
            MemoryStream ms = new MemoryStream(4096);
            BinaryWriter bw = new BinaryWriter(ms);

            foreach (var cls in m_classes.OrderBy(c => c.Key))
            {
                foreach (var part in cls.Value.Parts.OrderBy(p => p.Key))
                {
                    foreach (var state in part.Value.States.OrderBy(s => s.Key))
                    {
                        state.Value.Properties.Sort(Comparer<StyleProperty>.Create(
                            (p1, p2) =>
                            {
                                return p1.Header.nameID.CompareTo(p2.Header.nameID);
                            }));

                        foreach (var prop in state.Value.Properties)
                        {
                            PropertyStream.WriteProperty(bw, prop);
                        }
                    }
                }
            }

            // lang id
            ushort lid = ResourceAccess.GetFirstLanguageId(moduleHandle, "VARIANT", "NORMAL");
            byte[] data = ms.ToArray();
            return Win32Api.UpdateResource(updateHandle, "VARIANT", "NORMAL", lid, data, (uint)data.Length);
        }

        private void SaveStringTable(IntPtr moduleHandle, IntPtr updateHandle, bool makeStandalone)
        {
            if (m_stringTable.Count == 0)
            {
                return;
            }

            if (makeStandalone)
            {
                // To make a .msstyle relocatable, we need to remove the MUI (Multilingual UI) resource
                // entries and store the string table that was in the .mui's in the .msstyle itself.
                // This allows us to apply the style from anywhere, not restricted to its folder and the
                // accompanying mui resources.

                // Disadvantage: The style might work for only one script, because the string
                // table defines (among other unimportant things) the fonts. If a latin user creates a theme
                // this way, it might not work for on a chinese windows because the visual style might reference
                // latin-only fonts, so chinese texts won't render.


                // Deleting the MUI is tricky:
                // - If we attempt to delete a non-existing resource, all subsequent resource calls
                //   will fail with ERROR_INTERNAL_ERROR.
                // - If we attempt to delete a resource from a MUI file, we get ERROR_NOT_SUPPORTED.
                // - If we attempt to delete via LANG_NEUTRAL we get ERROR_INVALID_PARAMETER.

                ushort langId = ResourceAccess.GetFirstLanguageId(moduleHandle, "MUI", 1);
                if (langId != 0xFFFF)
                {
                    if (!Win32Api.UpdateResource(updateHandle, "MUI", 1, langId, null, 0))
                    {
                        int err = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
                        throw new Exception($"Deleting MUI for lang {langId} failed with error '{err}'!");
                    }
                }

                // Without the MUI, all updates should go into the .msstyles file rather.
                foreach (var table in m_stringTables)
                {
                    ResourceAccess.StringTable.Update(moduleHandle, updateHandle, table.Value, (ushort)table.Key);
                }
            }
            else
            {
                // We can't update the string table if it's in the .mui file. But since 
                // msstyleEditor doesn't allow editing the table, it not an immediate problem.
            }
        }

        private bool SaveAMap(IntPtr moduleHandle, IntPtr updateHandle)
        {
            if (m_animations.Count == 0 && m_timingFunctions.Count == 0)
                return true;

            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            foreach (var item in m_timingFunctions)
            {
                item.Write(bw);
            }
            foreach (var item in m_animations)
            {
                item.Write(bw);
            }

            ushort lid = ResourceAccess.GetFirstLanguageId(moduleHandle, "AMAP", "AMAP");
            if (lid == 0xFFFF)
            {
                lid = m_resourceLanguage;
            }
            byte[] data = ms.ToArray();
            return Win32Api.UpdateResource(updateHandle, "AMAP", "AMAP", lid, data, (uint)data.Length);

        }

        private bool SaveClassmap(IntPtr moduleHandle, IntPtr updateHandle)
        {
            // Preserve the original CMAP and append only new classes.
            // The original CMAP contains empty entries that act as separators;
            // LoadClassmap assigns sequential IDs skipping empties, and the
            // VARIANT properties reference those same IDs. We must not
            // reorder or compact the original entries.
            var ms = new MemoryStream();

            // Write original CMAP bytes as-is
            if (m_originalCmap != null)
            {
                ms.Write(m_originalCmap, 0, m_originalCmap.Length);
            }

            // Append new classes (those with IDs beyond what was loaded from original CMAP)
            var newClasses = m_classes
                .Where(kv => kv.Key >= m_originalClassCount)
                .OrderBy(kv => kv.Key);

            foreach (var kv in newClasses)
            {
                byte[] encoded = Encoding.Unicode.GetBytes(kv.Value.ClassName);
                ms.Write(encoded, 0, encoded.Length);
                ms.WriteByte(0); // null terminator (2 bytes for UTF-16)
                ms.WriteByte(0);
            }

            ushort lid = ResourceAccess.GetFirstLanguageId(moduleHandle, "CMAP", "CMAP");
            if (lid == 0xFFFF)
            {
                lid = m_resourceLanguage;
            }
            byte[] data = ms.ToArray();
            if (!Win32Api.UpdateResource(updateHandle, "CMAP", "CMAP", lid, data, (uint)data.Length))
                return false;

            // Update BCMAP: preserve original entries and append -1 for each new class.
            // BCMAP has one int32 per class (except the last 3 special classes).
            if (m_originalBcmap != null && newClasses.Any())
            {
                var bcms = new MemoryStream();
                bcms.Write(m_originalBcmap, 0, m_originalBcmap.Length);
                foreach (var kv in newClasses)
                {
                    // -1 means no parent class
                    bcms.Write(BitConverter.GetBytes(-1), 0, 4);
                }
                ushort bcmapLid = ResourceAccess.GetFirstLanguageId(moduleHandle, "BCMAP", "BCMAP");
                if (bcmapLid == 0xFFFF)
                    bcmapLid = m_resourceLanguage;
                byte[] bcmapData = bcms.ToArray();
                if (!Win32Api.UpdateResource(updateHandle, "BCMAP", "BCMAP", bcmapLid, bcmapData, (uint)bcmapData.Length))
                    return false;
            }

            return true;
        }
        public void Load(string file)
        {
            m_moduleHandle = Win32Api.LoadLibraryEx(file, IntPtr.Zero, Win32Api.LoadLibraryFlags.LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE);
            if (m_moduleHandle == IntPtr.Zero)
            {
                throw new Exception("Couldn't open file as PE resource!");
            }

            byte[] cmap = ResourceAccess.GetResource(m_moduleHandle, "CMAP", "CMAP");
            if (cmap == null)
            {
                // No CMAP — try XP format (TEXTFILE resources)
                byte[] themesIni = ResourceAccess.GetResource(m_moduleHandle, "TEXTFILE", "THEMES_INI");
                if (themesIni == null)
                {
                    throw new Exception("Style contains no class map and no THEMES_INI!");
                }

                LoadXp(themesIni);
                m_stylePath = file;
                return;
            }

            LoadClassmap(cmap);

            // Load BCMAP (Base Class Map) for class inheritance.
            // BCMAP is an int32[] where each entry is the parent classID or -1.
            // It has exactly (CMAP non-empty count - 3) entries.
            m_originalBcmap = ResourceAccess.GetResource(m_moduleHandle, "BCMAP", "BCMAP");

            // With the class map in place, we can reason about the platform.
            m_platform = DeterminePlatform();


            // There is no AMAP before Win8
            if (m_platform > Platform.Win7)
            {
                byte[] amap = ResourceAccess.GetResource(m_moduleHandle, "AMAP", "AMAP");
                if (amap == null)
                {
                    throw new Exception("Style contains no animation map!");
                }

                LoadAMap(amap);
            }


            BuildPropertyTree(m_platform);

            m_resourceLanguage = ResourceAccess.GetFirstLanguageId(m_moduleHandle, "VARIANT", "NORMAL");
            byte[] pmap = ResourceAccess.GetResource(m_moduleHandle, "VARIANT", "NORMAL");
            if (pmap == null)
            {
                throw new Exception("Style contains no properties!");
            }
            LoadProperties(pmap);


            // Try to get an overview of language resources.
            // Type: String Table, Name: 7 (typically style name & copyright)
            var l1 = ResourceAccess.GetAllLanguageIds(m_moduleHandle, "#" + Win32Api.RT_STRING, 7,
                Win32Api.EnumResourceFlags.RESOURCE_ENUM_MUI |
                Win32Api.EnumResourceFlags.RESOURCE_ENUM_LN);
            // Type: String Table, Name: 32 (typically font definitions)
            var l2 = ResourceAccess.GetAllLanguageIds(m_moduleHandle, "#" + Win32Api.RT_STRING, 32,
                Win32Api.EnumResourceFlags.RESOURCE_ENUM_MUI |
                Win32Api.EnumResourceFlags.RESOURCE_ENUM_LN);
            var langs = l1.Union(l2);

            // Load all tables for internal purposes.
            foreach (var lang in langs)
            {
                var table = new Dictionary<int, string>();
                ResourceAccess.StringTable.Load(m_moduleHandle, lang, table);
                m_stringTables[lang] = table;
            }

            // Get users preferred language for display purposes.
            // If we don't have it, choose any.
            int uiLang = System.Threading.Thread.CurrentThread.CurrentUICulture.LCID;
            if (!m_stringTables.TryGetValue(uiLang, out m_stringTable))
            {
                var kvp = m_stringTables.FirstOrDefault((t) => t.Value.Count > 0);
                m_stringTable = kvp.Value ?? new Dictionary<int, string>();
            }

            m_stylePath = file;
        }

        void LoadXp(byte[] themesIniData)
        {
            m_platform = Platform.WinXP;
            m_xpColorSchemes = XpStyleIniParser.ParseThemesIni(themesIniData);

            if (m_xpColorSchemes.Count == 0)
            {
                throw new Exception("No color schemes found in THEMES_INI!");
            }

            // Default to first scheme; caller can call LoadXpWithScheme() to switch
            LoadXpWithScheme(m_xpColorSchemes[0]);
        }

        /// <summary>
        /// Loads or reloads the XP style with the specified color scheme.
        /// Called initially from LoadXp() and can be called again when user selects a different scheme.
        /// </summary>
        public void LoadXpWithScheme(XpColorScheme scheme)
        {
            m_xpColorScheme = scheme;
            m_classes.Clear();
            m_xpImageMap.Clear();
            m_numProps = 0;

            byte[] iniData = ResourceAccess.GetResource(m_moduleHandle, "TEXTFILE", scheme.IniResourceName);
            if (iniData == null)
            {
                throw new Exception($"Could not load INI resource '{scheme.IniResourceName}'!");
            }

            XpStyleIniParser.ParseSchemeIni(iniData, m_classes);

            // Build image map: scan all FILENAME properties and map file paths to resource names
            foreach (var cls in m_classes)
            {
                foreach (var part in cls.Value.Parts)
                {
                    foreach (var state in part.Value.States)
                    {
                        foreach (var prop in state.Value.Properties)
                        {
                            if (prop.Header.typeID == (int)IDENTIFIER.FILENAME)
                            {
                                string filePath = prop.GetValue() as string;
                                if (!string.IsNullOrEmpty(filePath))
                                {
                                    string resName = XpStyleIniParser.ImageFileToResourceName(filePath);
                                    m_xpImageMap[filePath] = resName;
                                }
                            }
                            ++m_numProps;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Loads a BITMAP resource from the PE for XP styles.
        /// RT_BITMAP resources contain raw DIB data (no BITMAPFILEHEADER).
        /// Converts to PNG bytes with proper 32bpp alpha support.
        /// </summary>
        public byte[] GetXpBitmapResource(string resourceName)
        {
            if (string.IsNullOrEmpty(resourceName))
                return null;

            // RT_BITMAP = "#2"
            byte[] dib = ResourceAccess.GetResource(m_moduleHandle, "#2", resourceName);
            if (dib == null)
                return null;

            return ImageConverter.ConvertDibToPng(dib);
        }

        void LoadClassmap(byte[] cmap)
        {
            m_originalCmap = (byte[])cmap.Clone();

            int first = 0;
            int numFound = 0;
            int totalEntries = 0;

            for (int i = 0; i < cmap.Length; i += 2)
            {
                if (cmap[i] == 0 &&
                    cmap[i + 1] == 0)
                {
                    totalEntries++;

                    // we found the terminator and
                    // have a non-empty string
                    if (i - first > 2)
                    {
                        StyleClass cls = new StyleClass();
                        cls.ClassId = numFound;
                        cls.ClassName = Encoding.Unicode.GetString(cmap, first, i - first);
                        m_classes[numFound] = cls;
                        numFound++;
                    }

                    first = i + 2;
                }
            }

            m_cmapTotalEntries = totalEntries;
            m_originalClassCount = numFound;
        }

        void LoadProperties(byte[] pmap)
        {
            int cursor = 0;

            while (cursor < pmap.Length - 4)
            {
                try
                {
                    StyleProperty prop = PropertyStream.ReadNextProperty(pmap, ref cursor);

                    //Debug.WriteLine("[N: {0}, T: {1}, C: {2}, P: {3}, S: {4}]", prop.Header.nameID, prop.Header.typeID, prop.Header.classID, prop.Header.partID, prop.Header.stateID);

                    StyleClass cls;
                    if (!m_classes.TryGetValue(prop.Header.classID, out cls))
                    {
                        throw new Exception("Found property with unknown class ID");
                    }

                    StylePart part;
                    if (!cls.Parts.TryGetValue(prop.Header.partID, out part))
                    {
                        part = new StylePart()
                        {
                            PartId = prop.Header.partID,
                            PartName = "Part " + prop.Header.partID
                        };

                        cls.Parts.Add(part.PartId, part);
                    }

                    StyleState state;
                    if (!part.States.TryGetValue(prop.Header.stateID, out state))
                    {
                        state = new StyleState()
                        {
                            StateId = prop.Header.stateID,
                            StateName = "State " + prop.Header.stateID
                        };

                        part.States.Add(state.StateId, state);
                    }

                    state.Properties.Add(prop);
                    ++m_numProps;
                }
                catch (Exception)
                {

                }
            }
        }

        void LoadAMap(byte[] amap)
        {
            int cursor = 0;

            while (cursor < amap.Length - 4)
            {
                PropertyHeader header = new PropertyHeader(amap, cursor);

                cursor += 32;
                if (header.nameID == (int)IDENTIFIER.TIMINGFUNCTION)
                {
                    m_timingFunctions.Add(new TimingFunction(amap, cursor, header));
                    cursor += 24; // 20 bytes struct size, 4 bytes padding (= aligns to 8 byte?)
                }
                else if (header.nameID == (int)IDENTIFIER.ANIMATION)
                {
                    m_animations.Add(new Animation(amap, ref cursor, header));
                }
                else
                {
                    throw new Exception($"Unknown AMAP name ID: {header.nameID}");
                }
            }
        }

        Platform DeterminePlatform()
        {
            bool foundDWMTouch = false;
            bool foundDWMPen = false;
            bool foundW8Taskband = false;
            bool foundVistaQueryBuilder = false;
            bool foundTaskBand2Light_Taskband2 = false;

            foreach (var cls in m_classes)
            {
                if (cls.Value.ClassName == "DWMTouch")
                {
                    foundDWMTouch = true; continue;
                }
                if (cls.Value.ClassName == "DWMPen")
                {
                    foundDWMPen = true; continue;
                }
                if (cls.Value.ClassName == "W8::TaskbandExtendedUI")
                {
                    foundW8Taskband = true; continue;
                }
                if (cls.Value.ClassName == "QueryBuilder")
                {
                    foundVistaQueryBuilder = true; continue;
                }
                if (cls.Value.ClassName == "DarkMode::TaskManager")
                {
                    foundTaskBand2Light_Taskband2 = true; continue;
                }
            }

            if (foundTaskBand2Light_Taskband2)
                return Platform.Win11;
            else if (foundW8Taskband)
                return Platform.Win81;
            else if (foundDWMTouch || foundDWMPen)
                return Platform.Win10;
            else if (foundVistaQueryBuilder)
                return Platform.Vista;
            else return Platform.Win7;
        }

        void BuildPropertyTree(Platform p)
        {
            foreach (var cls in m_classes)
            {
                var partList = VisualStyleParts.Find(cls.Value.ClassName, p);
                foreach (var partDescription in partList)
                {
                    StylePart part = new StylePart()
                    {
                        PartId = partDescription.Id,
                        PartName = partDescription.Name,
                    };
                    cls.Value.Parts.Add(part.PartId, part);

                    foreach (var stateDescription in partDescription.States)
                    {
                        StyleState state = new StyleState()
                        {
                            StateId = stateDescription.Value,
                            StateName = stateDescription.Name,
                        };
                        part.States.Add(state.StateId, state);
                    }
                }
            }
        }

        public StyleResource GetResourceFromProperty(StyleProperty prop)
        {
            switch (prop.Header.typeID)
            {
                case (int)IDENTIFIER.FILENAME:
                case (int)IDENTIFIER.FILENAME_LITE:
                    {
                        // XP: FILENAME value is a string path, load from BITMAP resource
                        if (m_platform == Platform.WinXP)
                        {
                            string filePath = prop.GetValue() as string;
                            if (string.IsNullOrEmpty(filePath))
                                return null;

                            string resName;
                            if (!m_xpImageMap.TryGetValue(filePath, out resName))
                                resName = XpStyleIniParser.ImageFileToResourceName(filePath);

                            byte[] data = GetXpBitmapResource(resName);
                            // Use the hash code as a pseudo resource ID for XP
                            int resId = resName.GetHashCode();
                            return new StyleResource(data, resId, StyleResourceType.Image);
                        }

                        byte[] imgData = ResourceAccess.GetResource(m_moduleHandle, "IMAGE", (uint)prop.Header.shortFlag);
                        return new StyleResource(imgData, prop.Header.shortFlag, StyleResourceType.Image);
                    }
                case (int)IDENTIFIER.DISKSTREAM:
                    {
                        byte[] data = ResourceAccess.GetResource(m_moduleHandle, "STREAM", (uint)prop.Header.shortFlag);
                        return new StyleResource(data, prop.Header.shortFlag, StyleResourceType.Atlas);
                    }
                default:
                    {
                        return null;
                    }
            }
        }

        public string GetQueuedResourceUpdate(int nameId, StyleResourceType type)
        {
            string path;
            var key = new StyleResource(null, nameId, type);
            if (m_resourceUpdates.TryGetValue(key, out path))
            {
                return path;
            }
            else return string.Empty;
        }


        public void QueueResourceUpdate(int nameId, StyleResourceType type, string pathToNew)
        {
            var key = new StyleResource(null, nameId, type);
            m_resourceUpdates[key] = pathToNew;
        }
    }
}
