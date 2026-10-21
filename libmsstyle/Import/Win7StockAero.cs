using System;
using System.IO;
using System.Linq;

namespace libmsstyle.Import
{
    /// <summary>
    /// Stock windows 7 aero.msstyles, compiled into libmsstyle so the importer only needs the
    /// custom style. VisualStyle loads through LoadLibraryEx, which wants a file, so the bytes
    /// are written to %TEMP% once and reused.
    /// </summary>
    public static class Win7StockAero
    {
        private const string ResourceName = "libmsstyle.Import.w7_aero.msstyles";

        public static string GetPath()
        {
            using (var s = typeof(Win7StockAero).Assembly.GetManifestResourceStream(ResourceName))
            {
                if (s == null) throw new FileNotFoundException("stock win7 aero isn't embedded in this libmsstyle build");
                var ms = new MemoryStream();
                s.CopyTo(ms);
                byte[] data = ms.ToArray();

                // hash in the name: a different embedded copy never reuses a stale file.
                string sha = ResourceInventory.Sha1Hex(data).Substring(0, 12);
                string dir = Path.Combine(Path.GetTempPath(), "msstyleEditor");
                string path = Path.Combine(dir, $"w7_aero-{sha}.msstyles");
                if (File.Exists(path) && File.ReadAllBytes(path).SequenceEqual(data)) return path;

                Directory.CreateDirectory(dir);
                string tmp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                File.WriteAllBytes(tmp, data);
                try
                {
                    if (File.Exists(path)) File.Delete(path);
                    File.Move(tmp, path);
                }
                catch (IOException)
                {
                    // another instance won the race and may have it mapped; its copy is identical.
                    File.Delete(tmp);
                    if (!File.Exists(path)) throw;
                }
                return path;
            }
        }
    }
}
