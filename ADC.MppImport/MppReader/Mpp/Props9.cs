using System.IO;
using ADC.MppImport.MppReader.Common;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Props reader for MPP9 (MS Project 2003) files.
    /// Ported from org.mpxj.mpp.Props9
    /// </summary>
    internal class Props9 : Props
    {
        public Props9(byte[] data)
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

                    int itemSize = reader.ReadInt32();
                    int itemKey = reader.ReadInt32();
                    reader.ReadInt32(); // unknown
                    availableBytes -= 12;

                    if (availableBytes < itemSize || itemSize < 1) break;

                    byte[] itemData = reader.ReadBytes(itemSize);
                    availableBytes -= itemSize;

                    m_map[itemKey] = itemData;
                    ++foundCount;

                    // Align to two byte boundary
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
