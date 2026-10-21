using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

[assembly: InternalsVisibleTo("libmsstyleTests")]
namespace libmsstyle
{
    class PropertyStreamUnknownBytesException : Exception
    {
        public PropertyStreamUnknownBytesException(string message) : base(message) { }
    }

    class PropertyStream
    {
        public static StyleProperty ReadNextProperty(byte[] data, ref int start)
        {
            int cursor = start;
            // records are 8-byte aligned but string/list payloads aren't padded in the
            // length field, so step over padding until something looks like a header.
            while (cursor + 32 <= data.Length && !new PropertyHeader(data, cursor).IsValid())
            {
                cursor++;
            }

            if (cursor + 32 > data.Length)
            {
                int skipped = data.Length - start;
                start = data.Length;
                throw new PropertyStreamUnknownBytesException($"no valid property header in the last {skipped} bytes");
            }

            PropertyHeader header = new PropertyHeader(data, cursor);
            if (cursor - start > 4)
            {
                int skippedAt = start;
                start = cursor;
                throw new PropertyStreamUnknownBytesException($"skipped {cursor - skippedAt} unknown bytes at offset {skippedAt}");
            }

            cursor += 32; // sizeof PropertyHeader


            StyleProperty prop = new StyleProperty(
                (PropertyHeader)header.Clone()
            );

            switch ((IDENTIFIER)header.typeID)
            {
                case IDENTIFIER.INTLIST:
                    {
                        int numInts = 0;
                        var list = new List<Int32>();
                        if (header.sizeInBytes != 0)
                        {
                            numInts = BitConverter.ToInt32(data, cursor);
                            cursor += sizeof(Int32);
                        }

                        for(int i = 0; i < numInts; ++i)
                        {
                            list.Add(BitConverter.ToInt32(data, cursor));
                            cursor += sizeof(Int32);
                        }

                        prop.SetValue(list);
                        break;
                    }
                case IDENTIFIER.COLORLIST:
                    {
                        int numColors = header.sizeInBytes / 4;
                        var list = new List<Color>(numColors);

                        for (int i = 0; i < numColors; ++i)
                        {
                            int colorref = BitConverter.ToInt32(data, cursor);
                            list.Add(Color.FromArgb(colorref)); // TODO: convert?
                            cursor += 4;
                        }

                        prop.SetValue(list);
                        break;
                    }
                case IDENTIFIER.STRING:
                    {
                        int numChars = header.sizeInBytes / 2;
                        string text = "";

                        for (int i = 0; i < numChars - 1; ++i) // dont need the NULL term.
                        {
                            char c = BitConverter.ToChar(data, cursor);
                            text += c;
                            cursor += 2;
                        }
                        cursor += 2; // still need to account for the NULL term.
                        prop.SetValue(text);
                        break;
                    }
                // 32 byte property, (header carries data) 
                case IDENTIFIER.FILENAME:
                case IDENTIFIER.FILENAME_LITE:
                case IDENTIFIER.DISKSTREAM:
                case IDENTIFIER.FONT:
                    {
                        prop.SetValue(header.shortFlag);
                        break;
                    }
                // 40 byte property (32 byte header + 4 byte int + 4 byte padding)
                case IDENTIFIER.INT:
                case IDENTIFIER.SIZE:
                case IDENTIFIER.ENUM:
                case IDENTIFIER.HIGHCONTRASTCOLORTYPE:
                    {
                        if (header.shortFlag == 0)
                        {
                            prop.SetValue(BitConverter.ToInt32(data, cursor));
                            cursor += 8;
                        }
                        else prop.SetValue((Int32)0);
                        break;
                    }
                // 40 byte property, (32 byte header + 4 byte bool + 4 byte padding)
                case IDENTIFIER.BOOLTYPE:
                    {
                        if (header.shortFlag == 0)
                        {
                            prop.SetValue(BitConverter.ToInt32(data, cursor) != 0);
                            cursor += 8;
                        }
                        else prop.SetValue(false);
                        break;
                    }
                // 40 byte property, (32 byte header + 4 byte color type + 4 byte padding)
                case IDENTIFIER.COLOR:
                    {
                        if (header.shortFlag == 0)
                        {
                            int colorref = BitConverter.ToInt32(data, cursor);
                            int r = (colorref >> 0) & 0xFF;
                            int g = (colorref >> 8) & 0xFF;
                            int b = (colorref >> 16) & 0xFF;
                            prop.SetValue(Color.FromArgb(r,g,b));
                            cursor += 8;
                        }
                        else prop.SetValue(Color.FromArgb(0, 0, 0));
                        break;
                    }
                // 40 byte property, (32 byte header + 8 byte position type)
                case IDENTIFIER.POSITION:
                    {
                        if (header.shortFlag == 0)
                        {
                            int x = BitConverter.ToInt32(data, cursor);
                            cursor += 4;
                            int y = BitConverter.ToInt32(data, cursor);
                            cursor += 4;
                            prop.SetValue(new Size(x, y));
                        }
                        else prop.SetValue(new Size(0, 0));
                        break;
                    }
                // 48 byte property, (32 byte header + 12 byte rect type)
                case IDENTIFIER.RECTTYPE:
                    {
                        if (header.shortFlag == 0)
                        {
                            // struct RECT format is LTRB (windef.h)
                            int l = BitConverter.ToInt32(data, cursor);
                            cursor += 4;
                            int t = BitConverter.ToInt32(data, cursor);
                            cursor += 4;
                            int r = BitConverter.ToInt32(data, cursor);
                            cursor += 4;
                            int b = BitConverter.ToInt32(data, cursor);
                            cursor += 4;
                            prop.SetValue(new Margins(l, t, r, b));
                        }
                        else prop.SetValue(new Margins(0, 0, 0, 0));
                        break;
                    }
                case IDENTIFIER.MARGINS:
                    {
                        // struct MARGINS format is LRTB (uxtheme.h)
                        if (header.shortFlag == 0)
                        {
                            int l = BitConverter.ToInt32(data, cursor);
                            cursor += 4;
                            int r = BitConverter.ToInt32(data, cursor);
                            cursor += 4;
                            int t = BitConverter.ToInt32(data, cursor);
                            cursor += 4;
                            int b = BitConverter.ToInt32(data, cursor);
                            cursor += 4;
                            prop.SetValue(new Margins(l, t, r, b));
                        }
                        else prop.SetValue(new Margins(0, 0, 0, 0));
                        break;
                    }
                default:
                    {
                        // Copy the data of known props to an opaque memory block.
                        if (header.shortFlag == 0)
                        {
                            int sizeOfPayload = header.sizeInBytes;
                            if (sizeOfPayload > 0)
                            {
                                if (sizeOfPayload > data.Length - cursor)
                                {
                                    throw new InvalidDataException($"opaque property (type {header.typeID}) at offset {cursor - 32} claims {sizeOfPayload} bytes, only {data.Length - cursor} left");
                                }
                                byte[] mem = new byte[sizeOfPayload];
                                Array.Copy(data, cursor, mem, 0, sizeOfPayload);
                                cursor += sizeOfPayload;
                                prop.SetValue(mem);
                            }
                        }
                        break;
                    }
            }

            // uxtheme steps align8(32 + size). Only eat the gap when it's zero padding, so a
            // file that packs records without padding still falls through to the header scan.
            int aligned = (cursor + 7) & ~7;
            if (aligned <= data.Length)
            {
                bool zeros = true;
                for (int i = cursor; i < aligned; ++i) zeros &= data[i] == 0;
                if (zeros) cursor = aligned;
            }

            start = cursor;
            return prop;
        }

        private static void PadToMultipleOf(BinaryWriter writer, int align)
        {
            while (writer.BaseStream.Position % align != 0)
            {
                writer.Write((byte)0);
            }
        }

        public static void WriteProperty(BinaryWriter writer, StyleProperty prop)
        {
            writer.Write(prop.Header.Serialize());

            switch ((IDENTIFIER)prop.Header.typeID)
            {
                case IDENTIFIER.INTLIST:
                    {
                        if(prop.Header.sizeInBytes != 0)
                        {
                            var list = prop.GetValue() as List<Int32>;
                            writer.Write(list.Count);
                            
                            foreach(var num in list)
                            {
                                writer.Write(num);
                            }
                        }

                        PadToMultipleOf(writer, 8);
                        break;
                    }
                case IDENTIFIER.COLORLIST:
                    {
                        if (prop.Header.sizeInBytes != 0)
                        {
                            var list = prop.GetValue() as List<Color>;

                            foreach (var col in list)
                            {
                                writer.Write(col.B);
                                writer.Write(col.G);
                                writer.Write(col.R);
                                writer.Write(col.A);
                            }
                        }

                        PadToMultipleOf(writer, 8);
                        break;
                    }
                case IDENTIFIER.STRING:
                    {
                        var str = prop.GetValue() as string;
                        writer.Write(Encoding.Unicode.GetBytes(str));
                        writer.Write((ushort)0); // null term.

                        PadToMultipleOf(writer, 8);
                        break;
                    }
                // 32 byte property, (header carries data) 
                case IDENTIFIER.FILENAME:
                case IDENTIFIER.FILENAME_LITE:
                case IDENTIFIER.DISKSTREAM:
                case IDENTIFIER.FONT:
                    {
                        // data is in the header
                        break;
                    }
                // 40 byte property (32 byte header + 4 byte int + 4 byte padding)
                case IDENTIFIER.INT:
                case IDENTIFIER.SIZE:
                case IDENTIFIER.ENUM:
                case IDENTIFIER.HIGHCONTRASTCOLORTYPE:
                    {
                        if (prop.Header.shortFlag == 0)
                        {
                            writer.Write((int)prop.GetValue());
                        }

                        PadToMultipleOf(writer, 8);
                        break;
                    }
                // 40 byte property, (32 byte header + 4 byte bool + 4 byte padding)
                case IDENTIFIER.BOOLTYPE:
                    {
                        if (prop.Header.shortFlag == 0)
                        {
                            int val = (bool)prop.GetValue() ? 1 : 0;
                            writer.Write(val);
                        }

                        PadToMultipleOf(writer, 8);
                        break;
                    }
                // 40 byte property, (32 byte header + 4 byte color type + 4 byte padding)
                case IDENTIFIER.COLOR:
                    {
                        if (prop.Header.shortFlag == 0)
                        {
                            Color col = (Color)prop.GetValue();
                            writer.Write(col.R);
                            writer.Write(col.G);
                            writer.Write(col.B);
                            writer.Write((byte)0);
                        }

                        PadToMultipleOf(writer, 8);
                        break;
                    }
                // 40 byte property, (32 byte header + 8 byte position type)
                case IDENTIFIER.POSITION:
                    {
                        if (prop.Header.shortFlag == 0)
                        {
                            Size s = (Size)prop.GetValue();
                            writer.Write(s.Width);
                            writer.Write(s.Height);
                        }

                        PadToMultipleOf(writer, 8);
                        break;
                    }
                // 48 byte property, (32 byte header + 12 byte rect type)
                case IDENTIFIER.RECTTYPE:
                    {
                        // struct RECT format is LTRB (windef.h)
                        if (prop.Header.shortFlag == 0)
                        {
                            Margins m = prop.GetValue() as Margins;
                            writer.Write(m.Left);
                            writer.Write(m.Top);
                            writer.Write(m.Right);
                            writer.Write(m.Bottom);
                        }

                        PadToMultipleOf(writer, 8);
                        break;
                    }
                case IDENTIFIER.MARGINS:
                    {
                        // struct MARGINS format is LRTB (uxtheme.h)
                        if (prop.Header.shortFlag == 0)
                        {
                            Margins m = prop.GetValue() as Margins;
                            writer.Write(m.Left);
                            writer.Write(m.Right);
                            writer.Write(m.Top);
                            writer.Write(m.Bottom);
                        }

                        PadToMultipleOf(writer, 8);
                        break;
                    }
                default:
                    {
                        if (prop.Header.shortFlag == 0)
                        {
                            byte[] data = prop.GetValue() as byte[];
                            writer.Write(data);
                        }

                        PadToMultipleOf(writer, 8);
                        break;
                    }
            }
        }
    }
}
