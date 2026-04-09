using libmsstyle;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace msstyleEditor
{
    class ThemeManager
    {
        // This ends up calling
        // rundll32.exe uxtheme.dll,#64 C:\Windows\resources\Themes\Aero\Aero.msstyles?NormalColor?NormalSize
        // Can this be? Manually calling fails with access denied except on aero.

        [DllImport("uxtheme.dll", EntryPoint = "#65", CharSet = CharSet.Unicode)]
        private static extern uint SetSystemVisualStyle(string themeFile, string colorName, string sizeName, uint unknownFlags);

        [DllImport("uxtheme.dll", EntryPoint = "GetCurrentThemeName", CharSet = CharSet.Unicode)]
        private static extern uint GetCurrentThemeName(
            StringBuilder themeFile, int maxNameChars,
            StringBuilder colorName, int maxColorChars,
            StringBuilder sizeName, int maxSizeChars);

        [DllImport("user32.dll")]
        private static extern int GetSysColor(int nIndex);

        [DllImport("user32.dll")]
        private static extern bool SetSysColors(int cElements, int[] lpaElements, int[] lpaRgbValues);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool SystemParametersInfo(uint uiAction, uint uiParam, ref NONCLIENTMETRICS pvParam, uint fWinIni);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool SystemParametersInfo(uint uiAction, uint uiParam, ref bool pvParam, uint fWinIni);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);

        private const uint SPI_GETNONCLIENTMETRICS = 0x0029;
        private const uint SPI_SETNONCLIENTMETRICS = 0x002A;
        private const uint SPI_GETFLATMENU = 0x1022;
        private const uint SPI_SETFLATMENU = 0x1023;
        private const uint SPIF_UPDATEINIFILE = 0x01;
        private const uint SPIF_SENDCHANGE = 0x02;

        private static readonly int NUM_SYS_COLORS = 31;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct LOGFONT
        {
            public int lfHeight;
            public int lfWidth;
            public int lfEscapement;
            public int lfOrientation;
            public int lfWeight;
            public byte lfItalic;
            public byte lfUnderline;
            public byte lfStrikeOut;
            public byte lfCharSet;
            public byte lfOutPrecision;
            public byte lfClipPrecision;
            public byte lfQuality;
            public byte lfPitchAndFamily;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string lfFaceName;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct NONCLIENTMETRICS
        {
            public int cbSize;
            public int iBorderWidth;
            public int iScrollWidth;
            public int iScrollHeight;
            public int iCaptionWidth;
            public int iCaptionHeight;
            public LOGFONT lfCaptionFont;
            public int iSmCaptionWidth;
            public int iSmCaptionHeight;
            public LOGFONT lfSmCaptionFont;
            public int iMenuWidth;
            public int iMenuHeight;
            public LOGFONT lfMenuFont;
            public LOGFONT lfStatusFont;
            public LOGFONT lfMessageFont;
            public int iPaddedBorderWidth;
        }

        private Random m_rng = new Random();
        private string m_prevTheme;
        private string m_prevColor;
        private string m_prevSize;
        private int[] m_savedSysColorIndices;
        private int[] m_savedSysColorValues;
        private NONCLIENTMETRICS m_savedNcm;
        private bool m_savedNcmValid;
        private bool m_savedFlatMenu;
        private bool m_savedFlatMenuValid;

        private bool m_themeInUse;
        public bool IsThemeInUse { get { return m_themeInUse; } }

        private string m_customTheme;
        public string Theme { get { return m_customTheme; } }

        public ThemeManager()
        {
            StringBuilder t = new StringBuilder(260);
            StringBuilder c = new StringBuilder(260);
            StringBuilder s = new StringBuilder(260);
            if(GetCurrentThemeName(t, 260, c, 260, s, 260) != 0)
            {
                throw new SystemException("GetCurrentThemeName() failed!");
            }

            m_prevTheme = t.ToString();
            m_prevColor = c.ToString();
            m_prevSize = s.ToString();
        }

        ~ThemeManager()
        {
            try
            {
                Rollback();
            }
            catch(Exception) { }
        }

        public void ApplyTheme(VisualStyle style)
        {
            // Applying a style from any folder in /Users/<currentuser>/ breaks Win8.1 badly. (Issue #9)
            // This means /Desktop/, /AppData/, /Temp/, etc. cannot be used; weird...
            // A writeable alternative that doesn't crash Win8.1 is anything inside /Users/Public/.
            var pubDoc = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            var path = String.Format("{0}\\tmp{1:D5}.msstyles", pubDoc, m_rng.Next(0, 10000));

            // Save original system metrics BEFORE switching the theme,
            // so we capture the truly original values and not partially-changed ones.
            SaveOriginalMetrics();

            style.Save(path, true);
            uint hr = SetSystemVisualStyle(path,"NormalColor", "NormalSize", 0);
            if (hr != 0)
            {
                // We typically get codes like this: 0x90004004. To big for a plain Win32 error, so it must 
                // be a HRESULT. But 0x9 would mean, the N (= NTSTATUS) bit is set and that its just a warning?
                // Furthermore, 0x4004 doesn't translate to a meaningful NTSTATUS.
                // 
                // HRESULT:     S R C N X fffffffffff cccccccccccccccc
                // NTSTATUS:    S S C N f fffffffffff cccccccccccccccc
                //
                // If we clear the N bit, we'd get 0x80004004 which is COM error E_ABORT. Much better.
                // So we assume we get regular HRESULTs and ignore the bits R C N X.
                uint hresult = hr & 0x87FFFFFF;
                string msg = Marshal.GetExceptionForHR(unchecked((int)hresult)).Message;

                // 0x90004004 => COM error E_ABORT => Malformed style, OS can't read it
                // 0x90070490 => WIN32 error ERROR_NOT_FOUND ("Element not found.") => Resource missing

                File.Delete(path);

                throw new SystemException($"Failed to apply the theme as the OS rejected it! Message:\r\n\r\n{msg}");
            }
            else
            {
                if (!String.IsNullOrEmpty(m_customTheme))
                {
                    try
                    {
                        File.Delete(m_customTheme);
                    }
                    catch (Exception) { }
                }

                m_customTheme = path;
                m_themeInUse = true;

                // Wait for the OS to finish applying the visual style.
                // SetSystemVisualStyle triggers an async theme change that
                // can override system colors set too early.
                Thread.Sleep(500);

                // Apply system metrics from the style
                ApplySysMetrics(style);
            }
        }

        public void Rollback()
        {
            if (String.IsNullOrEmpty(m_customTheme))
            {
                return;
            }

            // Restore system colors before switching back
            RestoreSysMetrics();

            if (SetSystemVisualStyle(m_prevTheme, m_prevColor, m_prevSize, 0) != 0)
            {
                throw new SystemException("Failed to switch back to the previous theme!");
            }

            try
            {
                Thread.Sleep(250); // the OS takes a while to switch visual style
                File.Delete(m_customTheme);
            }
            catch (Exception) { }

            m_themeInUse = false;
        }

        public bool GetActiveTheme(out string theme, out string color, out string size)
        {
            StringBuilder t = new StringBuilder(260);
            StringBuilder c = new StringBuilder(260);
            StringBuilder s = new StringBuilder(260);

            bool res = GetCurrentThemeName(t, 260, c, 260, s, 260) == 0;

            theme = t.ToString();
            color = c.ToString();
            size = s.ToString();
            return res;
        }

        private void SaveOriginalMetrics()
        {
            // Save current system colors
            m_savedSysColorIndices = new int[NUM_SYS_COLORS];
            m_savedSysColorValues = new int[NUM_SYS_COLORS];
            for (int i = 0; i < NUM_SYS_COLORS; i++)
            {
                m_savedSysColorIndices[i] = i;
                m_savedSysColorValues[i] = GetSysColor(i);
            }

            // Save current non-client metrics (sizes + fonts)
            m_savedNcm = new NONCLIENTMETRICS();
            m_savedNcm.cbSize = Marshal.SizeOf(typeof(NONCLIENTMETRICS));
            m_savedNcmValid = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, (uint)m_savedNcm.cbSize, ref m_savedNcm, 0);

            // Save current flat menu setting
            m_savedFlatMenu = false;
            m_savedFlatMenuValid = SystemParametersInfo(SPI_GETFLATMENU, 0, ref m_savedFlatMenu, 0);
        }

        private void ApplySysMetrics(VisualStyle style)
        {
            // Only read sysmetrics from the "sysmetrics" class.
            // Other classes (e.g. TaskBand2) reuse the same nameIDs for
            // part-specific properties that are NOT system-wide metrics.
            StyleClass sysMetricsClass = null;
            foreach (var cls in style.Classes)
            {
                if (cls.Value.ClassName.Equals("sysmetrics", StringComparison.OrdinalIgnoreCase))
                {
                    sysMetricsClass = cls.Value;
                    break;
                }
            }

            if (sysMetricsClass == null)
                return;

            var colorMap = new Dictionary<int, int>();
            bool ncmModified = false;
            bool? flatMenuValue = null;

            var ncm = m_savedNcm; // start from current values, override with style values

            foreach (var part in sysMetricsClass.Parts)
            {
                foreach (var state in part.Value.States)
                {
                    foreach (var prop in state.Value.Properties)
                    {
                        int nameID = prop.Header.nameID;
                        int typeID = prop.Header.typeID;

                        // System colors (1601-1631)
                        if (nameID >= (int)IDENTIFIER.FIRSTCOLOR && nameID <= (int)IDENTIFIER.LASTCOLOR
                            && typeID == (int)IDENTIFIER.COLOR)
                        {
                            int colorIndex = nameID - (int)IDENTIFIER.FIRSTCOLOR;
                            var c = (Color)prop.GetValue();
                            // Construct COLORREF (0x00BBGGRR) directly
                            colorMap[colorIndex] = c.R | (c.G << 8) | (c.B << 16);
                        }
                        // System sizes (1201-1210)
                        else if (nameID >= (int)IDENTIFIER.FIRSTSIZE && nameID <= (int)IDENTIFIER.LASTSIZE
                            && typeID == (int)IDENTIFIER.SIZE)
                        {
                            int val = (int)prop.GetValue();
                            switch ((IDENTIFIER)nameID)
                            {
                                case IDENTIFIER.SIZINGBORDERWIDTH: ncm.iBorderWidth = val; ncmModified = true; break;
                                case IDENTIFIER.SCROLLBARWIDTH: ncm.iScrollWidth = val; ncmModified = true; break;
                                case IDENTIFIER.SCROLLBARHEIGHT: ncm.iScrollHeight = val; ncmModified = true; break;
                                case IDENTIFIER.CAPTIONBARWIDTH: ncm.iCaptionWidth = val; ncmModified = true; break;
                                case IDENTIFIER.CAPTIONBARHEIGHT: ncm.iCaptionHeight = val; ncmModified = true; break;
                                case IDENTIFIER.SMCAPTIONBARWIDTH: ncm.iSmCaptionWidth = val; ncmModified = true; break;
                                case IDENTIFIER.SMCAPTIONBARHEIGHT: ncm.iSmCaptionHeight = val; ncmModified = true; break;
                                case IDENTIFIER.MENUBARWIDTH: ncm.iMenuWidth = val; ncmModified = true; break;
                                case IDENTIFIER.MENUBARHEIGHT: ncm.iMenuHeight = val; ncmModified = true; break;
                                case IDENTIFIER.PADDEDBORDERWIDTH: ncm.iPaddedBorderWidth = val; ncmModified = true; break;
                            }
                        }
                        // System fonts (801-806)
                        else if (nameID >= (int)IDENTIFIER.FIRSTFONT && nameID <= 806
                            && typeID == (int)IDENTIFIER.FONT)
                        {
                            int fontResId = prop.Header.shortFlag;
                            string fontStr;
                            if (style.PreferredStringTable != null
                                && style.PreferredStringTable.TryGetValue(fontResId, out fontStr))
                            {
                                LOGFONT lf = ParseFontString(fontStr);
                                switch ((IDENTIFIER)nameID)
                                {
                                    case IDENTIFIER.CAPTIONFONT: ncm.lfCaptionFont = lf; ncmModified = true; break;
                                    case IDENTIFIER.SMALLCAPTIONFONT: ncm.lfSmCaptionFont = lf; ncmModified = true; break;
                                    case IDENTIFIER.MENUFONT: ncm.lfMenuFont = lf; ncmModified = true; break;
                                    case IDENTIFIER.STATUSFONT: ncm.lfStatusFont = lf; ncmModified = true; break;
                                    case IDENTIFIER.MSGBOXFONT: ncm.lfMessageFont = lf; ncmModified = true; break;
                                    case IDENTIFIER.ICONTITLEFONT: break; // ICONTITLEFONT is set separately via SPI_SETICONTITLELOGFONT, skip for now
                                }
                            }
                        }
                        // Flat menus (1001)
                        else if (nameID == (int)IDENTIFIER.FLATMENUS
                            && typeID == (int)IDENTIFIER.BOOLTYPE)
                        {
                            flatMenuValue = (bool)prop.GetValue();
                        }
                    }
                }
            }

            // Apply non-client metrics FIRST (sizes + fonts).
            // SPIF_SENDCHANGE broadcasts WM_SETTINGCHANGE which can cause
            // the OS to re-apply colors, so colors must be set AFTER this.
            if (ncmModified)
            {
                ncm.cbSize = Marshal.SizeOf(typeof(NONCLIENTMETRICS));
                SystemParametersInfo(SPI_SETNONCLIENTMETRICS, (uint)ncm.cbSize, ref ncm, SPIF_SENDCHANGE);
            }

            // Apply flat menu setting.
            // SPI_SETFLATMENU uses uiParam for the value, not pvParam.
            if (flatMenuValue.HasValue)
            {
                SystemParametersInfo(SPI_SETFLATMENU, flatMenuValue.Value ? 1u : 0u, IntPtr.Zero, SPIF_SENDCHANGE);
            }

            // Apply system colors LAST so that WM_SETTINGCHANGE from
            // the above calls cannot override our color values.
            if (colorMap.Count > 0)
            {
                var indices = new int[colorMap.Count];
                var values = new int[colorMap.Count];
                int idx = 0;
                foreach (var kvp in colorMap)
                {
                    indices[idx] = kvp.Key;
                    values[idx] = kvp.Value;
                    idx++;
                }
                SetSysColors(indices.Length, indices, values);
            }
        }

        private void RestoreSysMetrics()
        {
            // Restore system colors
            if (m_savedSysColorIndices != null && m_savedSysColorValues != null)
            {
                SetSysColors(m_savedSysColorIndices.Length, m_savedSysColorIndices, m_savedSysColorValues);
                m_savedSysColorIndices = null;
                m_savedSysColorValues = null;
            }

            // Restore non-client metrics
            if (m_savedNcmValid)
            {
                m_savedNcm.cbSize = Marshal.SizeOf(typeof(NONCLIENTMETRICS));
                SystemParametersInfo(SPI_SETNONCLIENTMETRICS, (uint)m_savedNcm.cbSize, ref m_savedNcm, SPIF_SENDCHANGE);
                m_savedNcmValid = false;
            }

            // Restore flat menu setting
            if (m_savedFlatMenuValid)
            {
                SystemParametersInfo(SPI_SETFLATMENU, 0, ref m_savedFlatMenu, SPIF_SENDCHANGE);
                m_savedFlatMenuValid = false;
            }
        }

        private static LOGFONT ParseFontString(string fontStr)
        {
            // Format: "Name, Size, [bold], [italic], [underline], [quality:Name|N]"
            var lf = new LOGFONT();
            var parts = fontStr.Split(new char[] { ',' }, StringSplitOptions.None);

            if (parts.Length >= 1)
            {
                lf.lfFaceName = parts[0].Trim();
            }

            if (parts.Length >= 2)
            {
                int size;
                if (Int32.TryParse(parts[1].Trim(), out size))
                {
                    // Convert point size to LOGFONT lfHeight
                    // lfHeight = -(size * DPI / 72), assume 96 DPI
                    lf.lfHeight = -(size * 96 / 72);
                }
            }

            lf.lfWeight = 400; // FW_NORMAL default
            lf.lfCharSet = 1; // DEFAULT_CHARSET

            for (int i = 2; i < parts.Length; i++)
            {
                string p = parts[i].Trim().ToLower();
                if (p == "bold") lf.lfWeight = 700;
                else if (p == "italic") lf.lfItalic = 1;
                else if (p == "underline") lf.lfUnderline = 1;
                else if (p.StartsWith("quality:"))
                {
                    var qualityToken = p.Substring(8).Trim();
                    int q;
                    if (Int32.TryParse(qualityToken, out q))
                    {
                        lf.lfQuality = (byte)q;
                    }
                    else
                    {
                        switch (qualityToken.Replace(" ", string.Empty))
                        {
                            case "default":
                                lf.lfQuality = 0;
                                break;
                            case "draft":
                                lf.lfQuality = 1;
                                break;
                            case "proof":
                                lf.lfQuality = 2;
                                break;
                            case "nonantialiased":
                            case "non-antialiased":
                                lf.lfQuality = 3;
                                break;
                            case "antialiased":
                            case "anti-aliased":
                                lf.lfQuality = 4;
                                break;
                            case "cleartype":
                                lf.lfQuality = 5;
                                break;
                            case "cleartypenatural":
                            case "cleartype-natural":
                                lf.lfQuality = 6;
                                break;
                        }
                    }
                }
            }

            return lf;
        }
    }
}
