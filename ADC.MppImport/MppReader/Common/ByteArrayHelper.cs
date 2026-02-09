using System;
using System.Text;

namespace ADC.MppImport.MppReader.Common
{
    /// <summary>
    /// Helper methods for working with byte arrays.
    /// Ported from org.mpxj.common.ByteArrayHelper
    /// </summary>
    public static class ByteArrayHelper
    {
        public static int GetShort(byte[] data, int offset)
        {
            int result = 0;
            int i = offset;
            for (int shiftBy = 0; shiftBy < 16; shiftBy += 8)
            {
                result |= ((data[i] & 0xff)) << shiftBy;
                ++i;
            }
            return result;
        }

        public static int GetInt(byte[] data, int offset)
        {
            int result = 0;
            int i = offset;
            for (int shiftBy = 0; shiftBy < 32; shiftBy += 8)
            {
                result |= ((data[i] & 0xff)) << shiftBy;
                ++i;
            }
            return result;
        }

        public static long GetLong(byte[] data, int offset)
        {
            long result = 0;
            int i = offset;
            for (int shiftBy = 0; shiftBy < 64; shiftBy += 8)
            {
                result |= ((long)(data[i] & 0xff)) << shiftBy;
                ++i;
            }
            return result;
        }

        public static string Hexdump(byte[] buffer, bool ascii)
        {
            int length = buffer != null ? buffer.Length : 0;
            return Hexdump(buffer, 0, length, ascii);
        }

        public static string Hexdump(byte[] buffer, int offset, int length, bool ascii)
        {
            if (buffer == null) return "";

            var sb = new StringBuilder();
            int count = offset + length;
            if (count > buffer.Length) count = buffer.Length;

            for (int loop = offset; loop < count; loop++)
            {
                if (sb.Length != 0) sb.Append(" ");
                sb.Append(HexDigits[(buffer[loop] & 0xF0) >> 4]);
                sb.Append(HexDigits[buffer[loop] & 0x0F]);
            }

            if (ascii)
            {
                sb.Append("   ");
                for (int loop = offset; loop < count; loop++)
                {
                    char c = (char)buffer[loop];
                    if (c > 200 || c < 27) c = ' ';
                    sb.Append(c);
                }
            }

            return sb.ToString();
        }

        private static readonly char[] HexDigits =
        {
            '0', '1', '2', '3', '4', '5', '6', '7',
            '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'
        };
    }
}
