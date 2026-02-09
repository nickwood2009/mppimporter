using System;
using System.Collections.Generic;
using System.IO;
using ADC.MppImport.MppReader.Common;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Reads variable-length data blocks associated with VarMeta metadata.
    /// Ported from org.mpxj.mpp.Var2Data
    /// </summary>
    internal class Var2Data
    {
        private readonly SortedDictionary<int, byte[]> m_map = new SortedDictionary<int, byte[]>();
        private readonly IVarMeta m_meta;

        public Var2Data(IVarMeta meta, byte[] buffer)
        {
            m_meta = meta;

            foreach (int itemOffset in meta.Offsets)
            {
                if (itemOffset < 0 || itemOffset + 4 > buffer.Length)
                    continue;

                int size = ByteArrayHelper.GetInt(buffer, itemOffset);
                if (size < 0 || itemOffset + 4 + size > buffer.Length)
                    continue;

                byte[] data = new byte[size];
                Array.Copy(buffer, itemOffset + 4, data, 0, size);
                m_map[itemOffset] = data;
            }
        }

        public byte[] GetByteArray(int? offset)
        {
            if (offset.HasValue && m_map.TryGetValue(offset.Value, out byte[] result))
                return result;
            return null;
        }

        public byte[] GetByteArray(int id, int type)
        {
            return GetByteArray(m_meta.GetOffset(id, type));
        }

        public string GetUnicodeString(int? offset)
        {
            if (!offset.HasValue) return null;
            byte[] value = GetByteArray(offset);
            if (value != null)
                return MppUtility.GetUnicodeString(value, 0);
            return null;
        }

        public string GetUnicodeString(int id, int type)
        {
            return GetUnicodeString(m_meta.GetOffset(id, type));
        }

        public DateTime? GetTimestamp(int id, int type)
        {
            int? offset = m_meta.GetOffset(id, type);
            if (!offset.HasValue) return null;
            byte[] value = GetByteArray(offset);
            if (value != null && value.Length >= 4)
                return MppUtility.GetTimestamp(value, 0);
            return null;
        }

        public string GetString(int? offset)
        {
            if (!offset.HasValue) return null;
            byte[] value = GetByteArray(offset);
            if (value != null)
                return MppUtility.GetString(value, 0);
            return null;
        }

        public string GetString(int id, int type)
        {
            return GetString(m_meta.GetOffset(id, type));
        }

        public int GetShort(int id, int type)
        {
            int? offset = m_meta.GetOffset(id, type);
            if (!offset.HasValue) return 0;
            byte[] value = GetByteArray(offset);
            if (value != null && value.Length >= 2)
                return ByteArrayHelper.GetShort(value, 0);
            return 0;
        }

        public int GetByte(int id, int type)
        {
            int? offset = m_meta.GetOffset(id, type);
            if (!offset.HasValue) return 0;
            byte[] value = GetByteArray(offset);
            if (value != null && value.Length > 0)
                return MppUtility.GetByte(value, 0);
            return 0;
        }

        public int GetInt(int id, int type)
        {
            int? offset = m_meta.GetOffset(id, type);
            if (!offset.HasValue) return 0;
            byte[] value = GetByteArray(offset);
            if (value != null && value.Length >= 4)
                return ByteArrayHelper.GetInt(value, 0);
            return 0;
        }

        public int GetInt(int id, int dataOffset, int type)
        {
            int? metaOffset = m_meta.GetOffset(id, type);
            if (!metaOffset.HasValue) return 0;
            byte[] value = GetByteArray(metaOffset);
            if (value != null && value.Length >= dataOffset + 4)
                return ByteArrayHelper.GetInt(value, dataOffset);
            return 0;
        }

        public long GetLong(int id, int type)
        {
            int? offset = m_meta.GetOffset(id, type);
            if (!offset.HasValue) return 0;
            byte[] value = GetByteArray(offset);
            if (value != null && value.Length >= 8)
                return ByteArrayHelper.GetLong(value, 0);
            return 0;
        }

        public double GetDouble(int id, int type)
        {
            double result = BitConverter.Int64BitsToDouble(GetLong(id, type));
            if (double.IsNaN(result)) result = 0;
            return result;
        }

        public IVarMeta VarMeta => m_meta;
    }
}
