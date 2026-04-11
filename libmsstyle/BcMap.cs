using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace libmsstyle
{
    /// <summary>
    /// Represents a single BCMAP entry containing a parent class index.
    /// BCMAP stores parent indices for class inheritance (-1 means no parent).
    /// </summary>
    public class BcMapEntry
    {
        /// <summary>
        /// Parent index in BCMAP space, or -1 if no parent.
        /// This is NOT the class ID - it's an index into the BCMAP array.
        /// </summary>
        public int ParentIndex { get; set; }

        public BcMapEntry(int parentIndex)
        {
            ParentIndex = parentIndex;
        }
    }

    /// <summary>
    /// Complete BCMAP (Base Class Map) structure.
    /// BCMAP defines class inheritance relationships via parent indices.
    /// </summary>
    public class BcMap
    {
        // Modern format has count field, legacy format doesn't
        public bool HasCountField { get; set; }

        // Parent indices for each class that has inheritance info
        // Index 0 of this list corresponds to the first class with BCMAP data
        public List<BcMapEntry> Entries { get; set; }

        // Number of "special" classes that come before BCMAP entries
        // Typically 4 for Vista+: document, sizevariant.NormalSize, sizevariant.Default, colorvariant.NormalColor
        public int SpecialClassCount { get; set; }

        public BcMap()
        {
            Entries = new List<BcMapEntry>();
            SpecialClassCount = 4;
        }

        /// <summary>
        /// Serializes the BCMAP to its binary format.
        /// Format: [int32 count]? + int32 parent indices + padding for 8-byte alignment
        /// </summary>
        public byte[] Serialize()
        {
            var ms = new MemoryStream();

            // Write count field if format requires it
            if (HasCountField)
            {
                byte[] countBytes = BitConverter.GetBytes(Entries.Count);
                ms.Write(countBytes, 0, 4);
            }

            // Write parent indices
            foreach (var entry in Entries)
            {
                byte[] indexBytes = BitConverter.GetBytes(entry.ParentIndex);
                ms.Write(indexBytes, 0, 4);
            }

            // Add padding for 8-byte alignment (WSB requirement)
            byte[] data = ms.ToArray();
            int paddingNeeded = (8 - (data.Length % 8)) % 8;
            if (paddingNeeded > 0)
            {
                Array.Resize(ref data, data.Length + paddingNeeded);
            }

            return data;
        }

        /// <summary>
        /// Converts a class ID to its corresponding BCMAP entry index.
        /// Class ID -> BCMAP index accounting for special classes.
        /// </summary>
        public int ClassIdToBcMapIndex(int classId)
        {
            return classId - SpecialClassCount;
        }

        /// <summary>
        /// Converts a BCMAP entry index to its corresponding class ID.
        /// BCMAP index -> Class ID accounting for special classes.
        /// </summary>
        public int BcMapIndexToClassId(int bcMapIndex)
        {
            return bcMapIndex + SpecialClassCount;
        }

        /// <summary>
        /// Converts a base class ID to a parent index for storage in BCMAP.
        /// </summary>
        public int BaseClassIdToParentIndex(int baseClassId)
        {
            if (baseClassId < 0)
                return -1;
            return ClassIdToBcMapIndex(baseClassId);
        }
    }
}
