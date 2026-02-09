using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ADC.MppImport.MppReader.Common;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Interface for VarMeta implementations.
    /// Ported from org.mpxj.mpp.VarMeta
    /// </summary>
    internal interface IVarMeta
    {
        int ItemCount { get; }
        int DataSize { get; }
        int[] GetUniqueIdentifierArray();
        HashSet<int> GetUniqueIdentifierSet();
        int? GetOffset(int id, int type);
        int[] Offsets { get; }
        HashSet<int> GetTypes(int id);
        bool ContainsKey(int key);
    }

    /// <summary>
    /// Base class for VarMeta implementations.
    /// Ported from org.mpxj.mpp.AbstractVarMeta
    /// </summary>
    internal abstract class AbstractVarMeta : IVarMeta
    {
        protected int m_itemCount;
        protected int m_dataSize;
        protected int[] m_offsets;
        protected readonly SortedDictionary<int, SortedDictionary<int, int>> m_table =
            new SortedDictionary<int, SortedDictionary<int, int>>();

        public int ItemCount => m_itemCount;
        public int DataSize => m_dataSize;
        public int[] Offsets => m_offsets;

        public int[] GetUniqueIdentifierArray()
        {
            return m_table.Keys.ToArray();
        }

        public HashSet<int> GetUniqueIdentifierSet()
        {
            return new HashSet<int>(m_table.Keys);
        }

        public int? GetOffset(int id, int type)
        {
            if (m_table.TryGetValue(id, out var map))
            {
                if (map.TryGetValue(type, out int offset))
                    return offset;
            }
            return null;
        }

        public HashSet<int> GetTypes(int id)
        {
            if (m_table.TryGetValue(id, out var map))
                return new HashSet<int>(map.Keys);
            return new HashSet<int>();
        }

        public bool ContainsKey(int key)
        {
            return m_table.ContainsKey(key);
        }

        protected void SetOffsets(int[] offsets)
        {
            m_offsets = offsets;
        }
    }

    /// <summary>
    /// VarMeta reader for MPP9 files.
    /// Ported from org.mpxj.mpp.VarMeta9
    /// </summary>
    internal class VarMeta9 : AbstractVarMeta
    {
        private const int MAGIC = unchecked((int)0xFADFADBA);

        public VarMeta9(byte[] data)
        {
            using (var ms = new MemoryStream(data))
            using (var reader = new BinaryReader(ms))
            {
                int magic = reader.ReadInt32();
                if (magic != MAGIC)
                    throw new IOException("Bad magic number: " + magic);

                reader.ReadInt32(); // unknown1
                m_itemCount = reader.ReadInt32();
                reader.ReadInt32(); // unknown2
                reader.ReadInt32(); // unknown3
                m_dataSize = reader.ReadInt32();

                int[] offsets = new int[m_itemCount];
                byte[] uniqueIDArray = new byte[4];

                for (int loop = 0; loop < m_itemCount; loop++)
                {
                    if (ms.Position + 8 > ms.Length) break;

                    // 3-byte unique ID
                    reader.Read(uniqueIDArray, 0, 3);
                    uniqueIDArray[3] = 0;
                    int uniqueID = ByteArrayHelper.GetInt(uniqueIDArray, 0);

                    int type = reader.ReadByte() & 0xFF;
                    int offset = reader.ReadInt32();

                    if (!m_table.TryGetValue(uniqueID, out var map))
                    {
                        map = new SortedDictionary<int, int>();
                        m_table[uniqueID] = map;
                    }
                    map[type] = offset;
                    offsets[loop] = offset;
                }

                Array.Sort(offsets);
                SetOffsets(offsets);
            }
        }
    }

    /// <summary>
    /// VarMeta reader for MPP12/MPP14 files.
    /// Ported from org.mpxj.mpp.VarMeta12
    /// </summary>
    internal class VarMeta12 : AbstractVarMeta
    {
        private const int MAGIC = unchecked((int)0xFADFADBA);

        public VarMeta12(byte[] data)
        {
            using (var ms = new MemoryStream(data))
            using (var reader = new BinaryReader(ms))
            {
                int magic = reader.ReadInt32();
                if (magic != 0 && magic != MAGIC)
                    throw new IOException("Bad magic number: " + magic);

                reader.ReadInt32(); // unknown1
                m_itemCount = reader.ReadInt32();
                reader.ReadInt32(); // unknown2
                reader.ReadInt32(); // unknown3
                m_dataSize = reader.ReadInt32();

                int[] offsets = new int[m_itemCount];

                for (int loop = 0; loop < m_itemCount; loop++)
                {
                    if (ms.Length - ms.Position < 12) break;

                    int uniqueID = reader.ReadInt32();
                    int offset = reader.ReadInt32();
                    int type = reader.ReadInt16() & 0xFFFF;
                    reader.ReadInt16(); // unknown 2 bytes

                    if (!m_table.TryGetValue(uniqueID, out var map))
                    {
                        map = new SortedDictionary<int, int>();
                        m_table[uniqueID] = map;
                    }
                    map[type] = offset;
                    offsets[loop] = offset;
                }

                Array.Sort(offsets);
                SetOffsets(offsets);
            }
        }
    }
}
