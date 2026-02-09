using System.IO;
using ADC.MppImport.MppReader.Common;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Props reader for MPP12 (MS Project 2007) files.
    /// Ported from org.mpxj.mpp.Props12
    /// </summary>
    internal class Props12 : Props
    {
        public Props12(byte[] data)
        {
            using (var ms = new MemoryStream(data))
            using (var reader = new BinaryReader(ms))
            {
                if (ms.Length < 16) return;

                byte[] header = reader.ReadBytes(16);
                int headerCount = ByteArrayHelper.GetShort(header, 12);
                int foundCount = 0;
                long availableBytes = ms.Length - ms.Position;

                while (foundCount < headerCount)
                {
                    if (availableBytes < 12) break;

                    int attrib1 = reader.ReadInt32();
                    int attrib2 = reader.ReadInt32();
                    reader.ReadInt32(); // unknown
                    availableBytes -= 12;

                    if (availableBytes < attrib1 || attrib1 < 1) break;

                    byte[] itemData = reader.ReadBytes(attrib1);
                    availableBytes -= attrib1;

                    m_map[attrib2] = itemData;
                    ++foundCount;

                    if (itemData.Length % 2 != 0 && ms.Position < ms.Length)
                    {
                        reader.ReadByte();
                        availableBytes--;
                    }
                }
            }
        }
    }
}
