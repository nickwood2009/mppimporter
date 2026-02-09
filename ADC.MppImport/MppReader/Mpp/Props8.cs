using System.IO;
using ADC.MppImport.MppReader.Common;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Props reader for MPP8 (MS Project 98/2000) files.
    /// Ported from org.mpxj.mpp.Props8
    /// </summary>
    internal class Props8 : Props
    {
        public bool Complete { get; private set; } = true;

        public Props8(byte[] data)
        {
            try
            {
                using (var ms = new MemoryStream(data))
                using (var reader = new BinaryReader(ms))
                {
                    reader.ReadInt32(); // File size
                    reader.ReadInt32(); // Repeat of file size
                    reader.ReadInt32(); // unknown
                    int count = reader.ReadInt16() & 0xFFFF;
                    reader.ReadInt16(); // unknown

                    for (int loop = 0; loop < count; loop++)
                    {
                        if (ms.Length - ms.Position < 12) break;

                        int attrib1 = reader.ReadInt32();

                        byte[] attrib = reader.ReadBytes(4);
                        int attrib2 = ByteArrayHelper.GetInt(attrib, 0);
                        int attrib3 = MppUtility.GetByte(attrib, 2);
                        int attrib5 = reader.ReadInt32();

                        int size;
                        if (attrib3 == 64)
                            size = attrib1;
                        else
                            size = attrib5;

                        if (attrib5 == 65536)
                            size = 4;

                        if (size <= 0)
                        {
                            Complete = false;
                            break;
                        }

                        if (ms.Position + size > ms.Length)
                        {
                            Complete = false;
                            break;
                        }

                        byte[] itemData = reader.ReadBytes(size);
                        m_map[attrib2] = itemData;

                        // Align to two byte boundary
                        if (itemData.Length % 2 != 0 && ms.Position < ms.Length)
                            reader.ReadByte();
                    }
                }
            }
            catch
            {
                Complete = false;
            }
        }
    }
}
