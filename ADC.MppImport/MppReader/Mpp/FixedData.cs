using System;
using System.IO;
using ADC.MppImport.MppReader.Common;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Reads FixedData blocks that contain fixed-size data records.
    /// Ported from org.mpxj.mpp.FixedData
    /// </summary>
    internal class FixedData
    {
        private readonly byte[][] m_array;
        private readonly int[] m_offset;

        /// <summary>
        /// Constructor using meta data to determine item positions and sizes.
        /// </summary>
        public FixedData(FixedMeta meta, byte[] buffer, int maxExpectedSize = 0, int minSize = 0)
        {
            int itemCount = meta.AdjustedItemCount;
            m_array = new byte[itemCount][];
            m_offset = new int[itemCount];

            for (int loop = 0; loop < itemCount; loop++)
            {
                byte[] metaData = meta.GetByteArrayValue(loop);
                if (metaData == null || metaData.Length < 8) continue;

                int itemOffset = ByteArrayHelper.GetInt(metaData, 4);
                if (itemOffset < 0 || itemOffset > buffer.Length) continue;

                int itemSize;
                if (loop + 1 == itemCount)
                {
                    itemSize = buffer.Length - itemOffset;
                }
                else
                {
                    byte[] nextMetaData = meta.GetByteArrayValue(loop + 1);
                    if (nextMetaData != null && nextMetaData.Length >= 8)
                    {
                        int nextItemOffset = ByteArrayHelper.GetInt(nextMetaData, 4);
                        itemSize = nextItemOffset - itemOffset;
                    }
                    else
                    {
                        itemSize = buffer.Length - itemOffset;
                    }
                }

                if (itemSize == 0) itemSize = minSize;

                int available = buffer.Length - itemOffset;
                if (itemSize < 0 || itemSize > available)
                {
                    itemSize = maxExpectedSize == 0 ? available : Math.Min(maxExpectedSize, available);
                }

                if (maxExpectedSize != 0 && itemSize > maxExpectedSize)
                    itemSize = maxExpectedSize;

                if (itemSize > 0)
                {
                    m_array[loop] = MppUtility.CloneSubArray(buffer, itemOffset, itemSize);
                    m_offset[loop] = itemOffset;
                }
            }
        }

        /// <summary>
        /// Constructor using a fixed item size.
        /// </summary>
        public FixedData(int itemSize, byte[] buffer)
        {
            int itemCount = buffer.Length / itemSize;
            m_array = new byte[itemCount][];
            m_offset = new int[itemCount];

            int offset = 0;
            for (int loop = 0; loop < itemCount; loop++)
            {
                m_offset[loop] = offset;
                int currentSize = Math.Min(itemSize, buffer.Length - offset);
                if (currentSize > 0)
                {
                    m_array[loop] = MppUtility.CloneSubArray(buffer, offset, currentSize);
                }
                offset += itemSize;
            }
        }

        /// <summary>
        /// Constructor using meta data with a specified item size override.
        /// </summary>
        public FixedData(FixedMeta meta, int itemSize, byte[] buffer)
        {
            int itemCount = meta.AdjustedItemCount;
            m_array = new byte[itemCount][];
            m_offset = new int[itemCount];

            for (int loop = 0; loop < itemCount; loop++)
            {
                byte[] metaData = meta.GetByteArrayValue(loop);
                if (metaData == null || metaData.Length < 8) continue;

                int itemOffset = ByteArrayHelper.GetInt(metaData, 4);
                if (itemOffset < 0 || itemOffset > buffer.Length) continue;

                int available = buffer.Length - itemOffset;
                int currentSize = Math.Min(itemSize < 0 ? available : itemSize, available);

                if (currentSize > 0)
                {
                    m_array[loop] = MppUtility.CloneSubArray(buffer, itemOffset, currentSize);
                    m_offset[loop] = itemOffset;
                }
            }
        }

        public byte[] GetByteArrayValue(int index)
        {
            if (index >= 0 && index < m_array.Length && m_array[index] != null)
                return m_array[index];
            return null;
        }

        public int ItemCount => m_array.Length;

        public bool IsValidOffset(int offset)
        {
            return offset >= 0 && offset < m_array.Length;
        }

        public int GetIndexFromOffset(int offset)
        {
            for (int loop = 0; loop < m_offset.Length; loop++)
            {
                if (m_offset[loop] == offset)
                    return loop;
            }
            return -1;
        }
    }
}
