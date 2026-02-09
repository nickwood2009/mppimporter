using System;
using System.IO;
using ADC.MppImport.MppReader.Common;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Reads FixedMeta blocks that describe the structure of FixedData blocks.
    /// Ported from org.mpxj.mpp.FixedMeta
    /// </summary>
    internal class FixedMeta
    {
        private readonly int m_itemCount;
        private readonly int m_adjustedItemCount;
        private readonly byte[][] m_array;

        private const int MAGIC = unchecked((int)0xFADFADBA);
        private const int HEADER_SIZE = 16;

        public FixedMeta(byte[] data, int itemSize)
        {
            using (var ms = new MemoryStream(data))
            using (var reader = new BinaryReader(ms))
            {
                int fileSize = data.Length;

                int magic = reader.ReadInt32();
                if (magic != MAGIC)
                    throw new IOException("Bad magic number: " + magic);

                reader.ReadInt32(); // unknown
                m_itemCount = reader.ReadInt32();
                reader.ReadInt32(); // unknown

                m_adjustedItemCount = (fileSize - HEADER_SIZE) / itemSize;
                m_array = new byte[m_adjustedItemCount][];

                for (int loop = 0; loop < m_adjustedItemCount; loop++)
                {
                    if (ms.Position + itemSize > ms.Length) break;
                    m_array[loop] = reader.ReadBytes(itemSize);
                }
            }
        }

        public int ItemCount => m_itemCount;
        public int AdjustedItemCount => m_adjustedItemCount;

        public byte[] GetByteArrayValue(int index)
        {
            if (index >= 0 && index < m_array.Length && m_array[index] != null)
                return m_array[index];
            return null;
        }
    }
}
