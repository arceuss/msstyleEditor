using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace libmsstyle.Import
{
    public class ResourceEntry
    {
        // "#123" for integer ids, otherwise the string name. Same convention as FindResource.
        public string Type { get; set; }
        public string Name { get; set; }
        public ushort Language { get; set; }
        public int Size { get; set; }
        public string Sha1 { get; set; }

        public string Key => Type + "/" + Name + "/" + Language;

        public bool TryGetIntName(out int id)
        {
            id = 0;
            return Name.StartsWith("#") && int.TryParse(Name.Substring(1), out id);
        }
    }

    /// <summary>
    /// Raw view of every resource (type/name/language) in a PE file, independent of the
    /// visual style parser. Used to prove that untouched resources survive a save untouched.
    /// </summary>
    public class ResourceInventory : IDisposable
    {
        private IntPtr m_module;
        private readonly Dictionary<string, byte[]> m_data = new Dictionary<string, byte[]>();
        public List<ResourceEntry> Entries { get; } = new List<ResourceEntry>();
        public string Path { get; }
        public string Machine { get; private set; }

        public ResourceInventory(string path)
        {
            Path = path;
            Machine = ReadMachine(path);
            m_module = Win32Api.LoadLibraryEx(path, IntPtr.Zero, Win32Api.LoadLibraryFlags.LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE);
            if (m_module == IntPtr.Zero)
                throw new IOException($"couldn't open \"{path}\" as a PE resource file (error {Marshal.GetLastWin32Error()}).");

            var types = new List<string>();
            // LN only: the MUI satellite isn't part of the file we write.
            Win32Api.EnumResourceTypesEx(m_module, (h, t, p) => { types.Add(PtrToResName(t)); return true; },
                IntPtr.Zero, Win32Api.EnumResourceFlags.RESOURCE_ENUM_LN, 0);

            foreach (var type in types)
            {
                var names = new List<string>();
                IntPtr typePtr = NameToPtr(type, out var typeHandle);
                try
                {
                    Win32Api.EnumResourceNamesEx(m_module, typePtr, (h, t, n, p) => { names.Add(PtrToResName(n)); return true; },
                        IntPtr.Zero, Win32Api.EnumResourceFlags.RESOURCE_ENUM_LN, 0);

                    foreach (var name in names)
                    {
                        IntPtr namePtr = NameToPtr(name, out var nameHandle);
                        try
                        {
                            var langs = new List<ushort>();
                            Win32Api.EnumResourceLanguagesEx(m_module, typePtr, namePtr, (h, t, n, l, p) => { langs.Add(l); return true; },
                                IntPtr.Zero, Win32Api.EnumResourceFlags.RESOURCE_ENUM_LN, 0);

                            foreach (var lang in langs)
                            {
                                byte[] bytes = Read(typePtr, namePtr, lang);
                                var e = new ResourceEntry
                                {
                                    Type = type,
                                    Name = name,
                                    Language = lang,
                                    Size = bytes?.Length ?? -1,
                                    Sha1 = bytes == null ? null : Sha1Hex(bytes)
                                };
                                Entries.Add(e);
                                if (bytes != null) m_data[e.Key] = bytes;
                            }
                        }
                        finally { if (nameHandle.IsAllocated) nameHandle.Free(); }
                    }
                }
                finally { if (typeHandle.IsAllocated) typeHandle.Free(); }
            }
        }

        public byte[] GetData(ResourceEntry e)
        {
            return m_data.TryGetValue(e.Key, out var d) ? d : null;
        }

        public IEnumerable<ResourceEntry> OfType(string type)
        {
            return Entries.Where(e => e.Type == type);
        }

        private byte[] Read(IntPtr type, IntPtr name, ushort lang)
        {
            IntPtr res = Win32Api.FindResourceEx(m_module, type, name, lang);
            if (res == IntPtr.Zero) return null;
            IntPtr h = Win32Api.LoadResource(m_module, res);
            IntPtr p = Win32Api.LockResource(h);
            uint size = Win32Api.SizeofResource(m_module, res);
            var buf = new byte[size];
            if (size > 0) Marshal.Copy(p, buf, 0, (int)size);
            return buf;
        }

        // IS_INTRESOURCE: pointer value <= 0xFFFF is an integer id.
        private static string PtrToResName(IntPtr p)
        {
            long v = p.ToInt64();
            if ((v >> 16) == 0) return "#" + v;
            return Marshal.PtrToStringUni(p);
        }

        private static IntPtr NameToPtr(string name, out GCHandle handle)
        {
            handle = default(GCHandle);
            if (name.StartsWith("#") && ushort.TryParse(name.Substring(1), out ushort id))
                return new IntPtr(id);
            var chars = (name + "\0").ToCharArray();
            handle = GCHandle.Alloc(chars, GCHandleType.Pinned);
            return handle.AddrOfPinnedObject();
        }

        public static string Sha1Hex(byte[] data)
        {
            using (var sha = SHA1.Create())
                return BitConverter.ToString(sha.ComputeHash(data)).Replace("-", "").ToLowerInvariant();
        }

        private static string ReadMachine(string path)
        {
            using (var fs = File.OpenRead(path))
            using (var br = new BinaryReader(fs))
            {
                if (fs.Length < 0x40 || br.ReadUInt16() != 0x5A4D) return "not-pe";
                fs.Seek(0x3C, SeekOrigin.Begin);
                int peOff = br.ReadInt32();
                if (peOff <= 0 || peOff + 6 > fs.Length) return "not-pe";
                fs.Seek(peOff, SeekOrigin.Begin);
                if (br.ReadUInt32() != 0x00004550) return "not-pe";
                ushort m = br.ReadUInt16();
                switch (m)
                {
                    case 0x014C: return "x86";
                    case 0x8664: return "x64";
                    case 0xAA64: return "arm64";
                    default: return "0x" + m.ToString("X4");
                }
            }
        }

        public void Dispose()
        {
            if (m_module != IntPtr.Zero)
            {
                Win32Api.FreeLibrary(m_module);
                m_module = IntPtr.Zero;
            }
        }
    }
}
