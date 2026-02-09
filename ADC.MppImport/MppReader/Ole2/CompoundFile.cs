using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ADC.MppImport.Ole2
{
    /// <summary>
    /// Minimal OLE2 Compound Binary File (MS-CFB) reader.
    /// Replaces the OpenMcdf NuGet dependency with zero external dependencies.
    /// Supports reading storages and streams from MPP files (v3 and v4).
    /// </summary>
    public class CompoundFile : IDisposable
    {
        private byte[] _data;
        private int _sectorSize;
        private int _miniSectorSize;
        private int _miniStreamCutoff;
        private int[] _fat;
        private int[] _miniFat;
        private byte[] _miniStream;
        private DirectoryEntry[] _dirEntries;

        private static readonly long SIGNATURE = BitConverter.ToInt64(
            new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }, 0);

        private const int ENDOFCHAIN = unchecked((int)0xFFFFFFFE);
        private const int FREESECT = unchecked((int)0xFFFFFFFF);

        public CFStorage RootStorage { get; private set; }

        public CompoundFile(Stream stream)
        {
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                _data = ms.ToArray();
            }
            Parse();
        }

        public CompoundFile(byte[] data)
        {
            _data = data;
            Parse();
        }

        private void Parse()
        {
            if (_data.Length < 512)
                throw new InvalidOperationException("File too small to be a compound file");

            if (BitConverter.ToInt64(_data, 0) != SIGNATURE)
                throw new InvalidOperationException("Invalid compound file signature");

            // Header fields (MS-CFB section 2.2)
            int sectorSizePower = BitConverter.ToUInt16(_data, 30);   // 9=512 (v3), 12=4096 (v4)
            int miniSectorSizePower = BitConverter.ToUInt16(_data, 32); // typically 6=64
            _sectorSize = 1 << sectorSizePower;
            _miniSectorSize = 1 << miniSectorSizePower;

            int totalFatSectors = BitConverter.ToInt32(_data, 44);
            int firstDirSector = BitConverter.ToInt32(_data, 48);
            _miniStreamCutoff = BitConverter.ToInt32(_data, 56);      // typically 4096
            int firstMiniFatSector = BitConverter.ToInt32(_data, 60);
            int totalMiniFatSectors = BitConverter.ToInt32(_data, 64);
            int firstDifatSector = BitConverter.ToInt32(_data, 68);
            int totalDifatSectors = BitConverter.ToInt32(_data, 72);

            BuildFat(totalFatSectors, firstDifatSector, totalDifatSectors);
            ReadDirectoryEntries(firstDirSector);
            ReadMiniFat(firstMiniFatSector, totalMiniFatSectors);
            ReadMiniStream();

            RootStorage = new CFStorage(this, 0);
        }

        private void BuildFat(int totalFatSectors, int firstDifatSector, int totalDifatSectors)
        {
            // Collect FAT sector IDs: first 109 from header DIFAT (offset 76), then from DIFAT chain
            var fatSectorIds = new List<int>();

            for (int i = 0; i < 109 && fatSectorIds.Count < totalFatSectors; i++)
            {
                int sid = BitConverter.ToInt32(_data, 76 + i * 4);
                if (sid < 0) break;
                fatSectorIds.Add(sid);
            }

            int difatSector = firstDifatSector;
            for (int d = 0; d < totalDifatSectors && difatSector >= 0 && difatSector != ENDOFCHAIN; d++)
            {
                int difatOffset = SectorOffset(difatSector);
                int entriesPerDifat = (_sectorSize / 4) - 1;
                for (int i = 0; i < entriesPerDifat && fatSectorIds.Count < totalFatSectors; i++)
                {
                    int sid = BitConverter.ToInt32(_data, difatOffset + i * 4);
                    if (sid < 0) break;
                    fatSectorIds.Add(sid);
                }
                difatSector = BitConverter.ToInt32(_data, difatOffset + entriesPerDifat * 4);
            }

            int entriesPerSector = _sectorSize / 4;
            _fat = new int[fatSectorIds.Count * entriesPerSector];
            for (int i = 0; i < fatSectorIds.Count; i++)
            {
                int offset = SectorOffset(fatSectorIds[i]);
                for (int j = 0; j < entriesPerSector; j++)
                {
                    int pos = offset + j * 4;
                    _fat[i * entriesPerSector + j] = pos + 4 <= _data.Length
                        ? BitConverter.ToInt32(_data, pos)
                        : FREESECT;
                }
            }
        }

        private void ReadDirectoryEntries(int firstDirSector)
        {
            byte[] dirData = ReadSectorChain(firstDirSector);
            int count = dirData.Length / 128;
            _dirEntries = new DirectoryEntry[count];
            for (int i = 0; i < count; i++)
                _dirEntries[i] = new DirectoryEntry(dirData, i * 128);
        }

        private void ReadMiniFat(int firstSector, int totalSectors)
        {
            if (firstSector == ENDOFCHAIN || firstSector < 0 || totalSectors == 0)
            {
                _miniFat = new int[0];
                return;
            }
            byte[] data = ReadSectorChain(firstSector);
            _miniFat = new int[data.Length / 4];
            for (int i = 0; i < _miniFat.Length; i++)
                _miniFat[i] = BitConverter.ToInt32(data, i * 4);
        }

        private void ReadMiniStream()
        {
            if (_dirEntries.Length == 0 || _dirEntries[0].StartSector < 0 ||
                _dirEntries[0].StartSector == ENDOFCHAIN || _dirEntries[0].StreamSize == 0)
            {
                _miniStream = new byte[0];
                return;
            }
            _miniStream = ReadSectorChain(_dirEntries[0].StartSector);
        }

        internal byte[] ReadStreamData(int startSector, long size)
        {
            if (size < _miniStreamCutoff && _miniFat.Length > 0)
                return ReadMiniChain(startSector, size);

            byte[] raw = ReadSectorChain(startSector);
            if (raw.Length > size)
            {
                byte[] trimmed = new byte[size];
                Array.Copy(raw, trimmed, (int)size);
                return trimmed;
            }
            return raw;
        }

        private byte[] ReadSectorChain(int startSector)
        {
            var sectors = new List<int>();
            int cur = startSector;
            int guard = _fat.Length + 1;
            while (cur >= 0 && cur != ENDOFCHAIN && guard-- > 0)
            {
                sectors.Add(cur);
                cur = cur < _fat.Length ? _fat[cur] : ENDOFCHAIN;
            }

            byte[] result = new byte[sectors.Count * _sectorSize];
            for (int i = 0; i < sectors.Count; i++)
            {
                int off = SectorOffset(sectors[i]);
                int len = Math.Min(_sectorSize, _data.Length - off);
                if (len > 0) Array.Copy(_data, off, result, i * _sectorSize, len);
            }
            return result;
        }

        private byte[] ReadMiniChain(int startSector, long size)
        {
            if (_miniStream.Length == 0) return new byte[0];

            var sectors = new List<int>();
            int cur = startSector;
            int guard = _miniFat.Length + 1;
            while (cur >= 0 && cur != ENDOFCHAIN && guard-- > 0)
            {
                sectors.Add(cur);
                cur = cur < _miniFat.Length ? _miniFat[cur] : ENDOFCHAIN;
            }

            byte[] result = new byte[size];
            int written = 0;
            for (int i = 0; i < sectors.Count && written < (int)size; i++)
            {
                int off = sectors[i] * _miniSectorSize;
                int len = Math.Min(_miniSectorSize, (int)size - written);
                if (off + len <= _miniStream.Length)
                    Array.Copy(_miniStream, off, result, written, len);
                written += len;
            }
            return result;
        }

        private int SectorOffset(int sectorId)
        {
            return (sectorId + 1) * _sectorSize;
        }

        internal DirectoryEntry GetEntry(int index)
        {
            return (index >= 0 && index < _dirEntries.Length) ? _dirEntries[index] : null;
        }

        internal List<int> GetChildIndices(int parentIndex)
        {
            var result = new List<int>();
            var entry = GetEntry(parentIndex);
            if (entry != null && entry.ChildSid >= 0 && entry.ChildSid != unchecked((int)0xFFFFFFFF))
                CollectTree(entry.ChildSid, result, new HashSet<int>());
            return result;
        }

        private void CollectTree(int index, List<int> result, HashSet<int> visited)
        {
            if (index < 0 || index >= _dirEntries.Length || index == unchecked((int)0xFFFFFFFF))
                return;
            if (!visited.Add(index)) return; // cycle guard

            var e = _dirEntries[index];
            CollectTree(e.LeftSiblingSid, result, visited);
            result.Add(index);
            CollectTree(e.RightSiblingSid, result, visited);
        }

        public void Dispose() { }
    }

    internal class DirectoryEntry
    {
        public string Name;
        public byte ObjectType; // 0=unknown, 1=storage, 2=stream, 5=root
        public int LeftSiblingSid;
        public int RightSiblingSid;
        public int ChildSid;
        public int StartSector;
        public long StreamSize;

        public DirectoryEntry(byte[] data, int offset)
        {
            int nameBytes = BitConverter.ToUInt16(data, offset + 64);
            Name = nameBytes > 2
                ? Encoding.Unicode.GetString(data, offset, nameBytes - 2)
                : "";

            ObjectType = data[offset + 66];
            LeftSiblingSid = BitConverter.ToInt32(data, offset + 68);
            RightSiblingSid = BitConverter.ToInt32(data, offset + 72);
            ChildSid = BitConverter.ToInt32(data, offset + 76);
            StartSector = BitConverter.ToInt32(data, offset + 116);
            StreamSize = BitConverter.ToInt64(data, offset + 120);
        }

        public bool IsStorage => ObjectType == 1 || ObjectType == 5;
        public bool IsStream => ObjectType == 2;
    }

    /// <summary>
    /// Base class for compound file entries (storages and streams).
    /// </summary>
    public class CFItem
    {
        internal CompoundFile _cf;
        internal int _dirIndex;

        internal CFItem(CompoundFile cf, int dirIndex)
        {
            _cf = cf;
            _dirIndex = dirIndex;
        }

        public string Name => _cf.GetEntry(_dirIndex)?.Name ?? "";

        public bool IsStorage
        {
            get
            {
                var e = _cf.GetEntry(_dirIndex);
                return e != null && e.IsStorage;
            }
        }
    }

    /// <summary>
    /// A storage (directory) entry within a compound file.
    /// </summary>
    public class CFStorage : CFItem
    {
        internal CFStorage(CompoundFile cf, int dirIndex) : base(cf, dirIndex) { }

        public CFStorage GetStorage(string name)
        {
            foreach (int ci in _cf.GetChildIndices(_dirIndex))
            {
                var e = _cf.GetEntry(ci);
                if (e != null && e.IsStorage &&
                    string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase))
                    return new CFStorage(_cf, ci);
            }
            throw new KeyNotFoundException("Storage not found: " + name);
        }

        public CFStream GetStream(string name)
        {
            foreach (int ci in _cf.GetChildIndices(_dirIndex))
            {
                var e = _cf.GetEntry(ci);
                if (e != null && e.IsStream &&
                    string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase))
                    return new CFStream(_cf, ci);
            }
            throw new KeyNotFoundException("Stream not found: " + name);
        }

        public void VisitEntries(Action<CFItem> action, bool recursive)
        {
            foreach (int ci in _cf.GetChildIndices(_dirIndex))
            {
                var e = _cf.GetEntry(ci);
                if (e == null) continue;

                CFItem item = e.IsStorage
                    ? (CFItem)new CFStorage(_cf, ci)
                    : new CFStream(_cf, ci);

                action(item);

                if (recursive && e.IsStorage)
                    ((CFStorage)item).VisitEntries(action, true);
            }
        }
    }

    /// <summary>
    /// A stream (data) entry within a compound file.
    /// </summary>
    public class CFStream : CFItem
    {
        internal CFStream(CompoundFile cf, int dirIndex) : base(cf, dirIndex) { }

        public byte[] GetData()
        {
            var e = _cf.GetEntry(_dirIndex);
            if (e == null || e.StreamSize == 0) return new byte[0];
            return _cf.ReadStreamData(e.StartSector, e.StreamSize);
        }
    }
}
