using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace libmsstyle
{
    public class ImageConverter
    {
        public static void PremultiplyAlpha(Stream instream, out Bitmap dst)
        {
            var src = new Bitmap(instream);
            // Premultiply the alpha channel.
            var data = src.LockBits(new Rectangle(0, 0, src.Width, src.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppPArgb);
            // Here, we lie. The data is actually premultiplied, not straight. If we don't lie,
            // the bitmap class will convert the data from PArgb back to Argb when saving as PNG.
            // It does so, because PNG has per specification always straight alpha.
            dst = new Bitmap(src.Width, src.Height, data.Stride, PixelFormat.Format32bppArgb, data.Scan0);
        }

        public static void PremulToStraightAlpha(Stream instream, out Bitmap dst) 
        {
            var src = new Bitmap(instream);
            // Bitmap `src` will have format Argb, because PNGs are always supposed to be Argb and there is no
            // field in a PNG image describing this. LockBits(Argb) will thus perform no conversion.
            var data = src.LockBits(new Rectangle(0, 0, src.Width, src.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            // This bitmap, we tell the truth about `data` having premultiplied alpha. When saving the
            // bitmap as PNG (done by the caller of this function), the conversion to straight alpha will happen.
            dst = new Bitmap(src.Width, src.Height, data.Stride, PixelFormat.Format32bppPArgb, data.Scan0);
        }

        /// <summary>
        /// Converts raw DIB data (from RT_BITMAP PE resources) to PNG byte array.
        /// Handles 32bpp with proper alpha preservation (GDI+ normally drops alpha on BMP load).
        /// </summary>
        public static byte[] ConvertDibToPng(byte[] dib)
        {
            short biBitCount = BitConverter.ToInt16(dib, 14);

            using (Bitmap bmp = biBitCount == 32 ? CreateArgbBitmapFromDib(dib) : CreateBitmapFromDib(dib))
            using (var ms = new MemoryStream())
            {
                bmp.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// Manually constructs a 32bpp ARGB Bitmap from raw DIB data.
        /// GDI+ loads 32bpp BMPs as Format32bppRgb, discarding alpha — this bypasses that.
        /// </summary>
        private static Bitmap CreateArgbBitmapFromDib(byte[] dib)
        {
            int biSize = BitConverter.ToInt32(dib, 0);
            int biWidth = BitConverter.ToInt32(dib, 4);
            int biHeight = BitConverter.ToInt32(dib, 8);
            int biCompression = BitConverter.ToInt32(dib, 16);

            bool bottomUp = biHeight > 0;
            int height = Math.Abs(biHeight);

            // For BI_BITFIELDS (3) with standard 40-byte header, 3 DWORD masks follow
            int extraBytes = 0;
            if (biCompression == 3 && biSize == 40)
                extraBytes = 12;

            int pixelOffset = biSize + extraBytes;
            int stride = biWidth * 4;

            var bmp = new Bitmap(biWidth, height, PixelFormat.Format32bppArgb);
            var bmpData = bmp.LockBits(new Rectangle(0, 0, biWidth, height),
                ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);

            for (int y = 0; y < height; y++)
            {
                int srcY = bottomUp ? (height - 1 - y) : y;
                int srcOff = pixelOffset + srcY * stride;
                IntPtr dstPtr = IntPtr.Add(bmpData.Scan0, y * bmpData.Stride);
                if (srcOff + stride <= dib.Length)
                    Marshal.Copy(dib, srcOff, dstPtr, stride);
            }

            bmp.UnlockBits(bmpData);
            return bmp;
        }

        /// <summary>
        /// Creates a Bitmap from non-32bpp DIB data using GDI+ with a prepended file header.
        /// Returns a deep copy that is independent of any stream.
        /// </summary>
        private static Bitmap CreateBitmapFromDib(byte[] dib)
        {
            byte[] bmpFile = PrependBitmapFileHeader(dib);
            using (var ms = new MemoryStream(bmpFile))
            using (var temp = new Bitmap(ms))
            {
                return new Bitmap(temp);
            }
        }

        private static byte[] PrependBitmapFileHeader(byte[] dib)
        {
            int biSize = BitConverter.ToInt32(dib, 0);
            short biBitCount = BitConverter.ToInt16(dib, 14);
            int biCompression = BitConverter.ToInt32(dib, 16);
            int biClrUsed = BitConverter.ToInt32(dib, 32);

            int colorTableSize = 0;
            if (biBitCount <= 8)
            {
                int colors = biClrUsed > 0 ? biClrUsed : (1 << biBitCount);
                colorTableSize = colors * 4;
            }
            else if (biCompression == 3 && biSize == 40)
            {
                colorTableSize = 12;
            }

            int fileHeaderSize = 14;
            int pixelDataOffset = fileHeaderSize + biSize + colorTableSize;

            byte[] bmpFile = new byte[fileHeaderSize + dib.Length];
            bmpFile[0] = 0x42; // 'B'
            bmpFile[1] = 0x4D; // 'M'
            Array.Copy(BitConverter.GetBytes(bmpFile.Length), 0, bmpFile, 2, 4);
            Array.Copy(BitConverter.GetBytes(pixelDataOffset), 0, bmpFile, 10, 4);
            Array.Copy(dib, 0, bmpFile, fileHeaderSize, dib.Length);

            return bmpFile;
        }
    }
}
