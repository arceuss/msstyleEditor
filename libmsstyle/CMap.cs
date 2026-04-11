using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace libmsstyle
{
    /// <summary>
    /// Represents a single CMAP entry containing a class name.
    /// CMAP stores UTF-16LE class name strings for the uxtheme class system.
    /// </summary>
    public class CMapEntry
    {
        public string ClassName { get; set; }

        public CMapEntry(string className)
        {
            ClassName = className;
        }
    }

    /// <summary>
    /// Complete CMAP (Class Map) structure.
    /// CMAP is a UTF-16LE string array with null terminators and alignment padding.
    /// </summary>
    public class CMap
    {
        // Alignment: 4 bytes on x86, 8 bytes on x64
        public int EntryAlignment { get; set; }

        // All class name entries in order
        public List<CMapEntry> Entries { get; set; }

        // Whether this CMAP was originally malformed and repaired
        public bool WasRepaired { get; set; }

        public CMap()
        {
            EntryAlignment = (IntPtr.Size >= 8) ? 8 : 4;
            Entries = new List<CMapEntry>();
        }

        /// <summary>
        /// Serializes the CMAP to its binary format.
        /// Format: [UTF16 string + null terminator + padding] repeated for each entry
        /// </summary>
        public byte[] Serialize()
        {
            var ms = new MemoryStream();

            // Filter out empty or whitespace-only entries to prevent blank lines in output
            var validEntries = Entries.FindAll(e => !string.IsNullOrWhiteSpace(e.ClassName));

            if (validEntries.Count > 0)
            {
                for (int i = 0; i < validEntries.Count; i++)
                {
                    // Don't add padding after the last entry (only the null terminator)
                    bool isLastEntry = (i == validEntries.Count - 1);
                    WriteCmapEntry(ms, validEntries[i].ClassName, EntryAlignment, !isLastEntry);
                }
            }
            else
            {
                // Empty CMAP - write just a null terminator
                ms.WriteByte(0);
                ms.WriteByte(0);
            }
            return ms.ToArray();
        }

        private static void WriteCmapEntry(Stream stream, string className, int alignment, bool addPadding)
        {
            // Write UTF-16LE string
            byte[] stringBytes = Encoding.Unicode.GetBytes(className);
            stream.Write(stringBytes, 0, stringBytes.Length);

            // Write null terminator (2 bytes for UTF-16)
            stream.WriteByte(0);
            stream.WriteByte(0);

            // Add padding to align (only for non-last entries)
            if (addPadding)
            {
                long currentPos = stream.Position;
                int paddingNeeded = (alignment - (int)(currentPos % alignment)) % alignment;
                for (int i = 0; i < paddingNeeded; i++)
                {
                    stream.WriteByte(0);
                }
            }
        }
    }
}
