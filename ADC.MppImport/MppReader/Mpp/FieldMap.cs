using System;
using System.Collections.Generic;
using ADC.MppImport.MppReader.Common;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Known task field indices from MPPTaskField.FIELD_ARRAY.
    /// The index is the lower 16 bits of the field type ID (upper 16 = 0x0B40 for tasks).
    /// </summary>
    internal enum TaskFieldIndex
    {
        Work = 0,
        BaselineWork = 1,
        ActualWork = 2,
        RemainingWork = 4,
        Cost = 5,
        BaselineCost = 6,
        ActualCost = 7,
        FixedCost = 8,
        RemainingCost = 10,
        Name = 14,
        Notes = 15,
        WBS = 16,
        ConstraintType = 17,
        ConstraintDate = 18,
        FreeSlack = 21,
        TotalSlack = 22,
        ID = 23,
        Milestone = 24,
        Priority = 25,
        BaselineDuration = 27,
        ActualDuration = 28,
        Duration = 29,
        RemainingDuration = 31,
        PercentComplete = 32,
        PercentWorkComplete = 33,
        Start = 35,
        Finish = 36,
        EarlyStart = 37,
        EarlyFinish = 38,
        LateStart = 39,
        LateFinish = 40,
        ActualStart = 41,
        ActualFinish = 42,
        BaselineStart = 43,
        BaselineFinish = 44,
        OutlineLevel = 85,
        UniqueID = 86,
        Created = 93,
        Resume = 99,
        Stop = 100,
        Type = 128,
        DurationUnits = 152,
        ParentTaskUniqueID = 160,
        FixedCostAccrual = 200,
        Deadline = 437,
        CalendarUniqueID = 401,
        Active = 1279,
    }

    /// <summary>
    /// Known resource field indices from MPPResourceField.FIELD_ARRAY.
    /// The index is the lower 16 bits of the field type ID (upper 16 = 0x0C400000 for resources).
    /// </summary>
    internal enum ResourceFieldIndex
    {
        ID = 0,
        Name = 1,
        Initials = 2,
        Group = 3,
        MaxUnits = 4,
        StandardRate = 6,
        OvertimeRate = 7,
        Code = 10,
        ActualCost = 11,
        Cost = 12,
        Work = 13,
        ActualWork = 14,
        CostPerUse = 18,
        AccrueAt = 19,
        Notes = 20,
        RemainingCost = 21,
        RemainingWork = 22,
        UniqueID = 27,
        PercentWorkComplete = 29,
        EmailAddress = 35,
        CreationDate = 726,
        CalendarUniqueID = 9990,
        Active = 9991,
    }

    /// <summary>
    /// Known assignment field indices from MPPAssignmentField.FIELD_ARRAY.
    /// </summary>
    internal enum AssignmentFieldIndex
    {
        UniqueID = 0,
        TaskUniqueID = 1,
        ResourceUniqueID = 2,
        Units = 7,
        Work = 8,
        ActualWork = 10,
        RemainingWork = 12,
        Start = 20,
        Finish = 21,
        Cost = 26,
        ActualCost = 28,
        PercentWorkComplete = 43,
    }

    /// <summary>
    /// Where a field's data is located.
    /// </summary>
    internal enum FieldLocation
    {
        FixedData,
        VarData,
        MetaData,
        Unknown
    }

    /// <summary>
    /// Represents the location of a single field within the data blocks.
    /// </summary>
    internal class FieldItem
    {
        public int FieldIndex { get; set; }
        public FieldLocation Location { get; set; }
        public int DataBlockIndex { get; set; }
        public int DataBlockOffset { get; set; }
        public int VarDataKey { get; set; }
        public int Category { get; set; }
    }

    /// <summary>
    /// Reads and parses field map data from Props to determine where each field
    /// lives in the fixed data and variable data blocks.
    /// Ported from org.mpxj.mpp.FieldMap
    /// </summary>
    internal class FieldMap
    {
        private readonly Dictionary<int, FieldItem> m_map = new Dictionary<int, FieldItem>();
        private readonly int[] m_maxFixedDataSize = new int[4];
        private readonly bool m_useTypeAsVarDataKey;

        // Props keys for field map data
        public const int TASK_FIELD_MAP_KEY = 131092;
        public const int TASK_FIELD_MAP_KEY2 = 50331668;
        public const int RESOURCE_FIELD_MAP_KEY = 131093;
        public const int RESOURCE_FIELD_MAP_KEY2 = 50331669;
        public const int ASSIGNMENT_FIELD_MAP_KEY = 131095;
        public const int ASSIGNMENT_FIELD_MAP_KEY2 = 50331671;

        // Field type prefixes (upper 16 bits)
        public const int TASK_FIELD_BASE = 0x0B400000;
        public const int RESOURCE_FIELD_BASE = 0x0C400000;
        public const int ASSIGNMENT_FIELD_BASE = 0x0B500000;

        public FieldMap(bool useTypeAsVarDataKey)
        {
            m_useTypeAsVarDataKey = useTypeAsVarDataKey;
        }

        /// <summary>
        /// Create a task field map from Props data.
        /// </summary>
        public void CreateTaskFieldMap(Props props)
        {
            byte[] data = props.GetByteArray(TASK_FIELD_MAP_KEY);
            if (data == null) data = props.GetByteArray(TASK_FIELD_MAP_KEY2);
            if (data != null) ParseFieldMap(data);
        }

        /// <summary>
        /// Create a resource field map from Props data.
        /// </summary>
        public void CreateResourceFieldMap(Props props)
        {
            byte[] data = props.GetByteArray(RESOURCE_FIELD_MAP_KEY);
            if (data == null) data = props.GetByteArray(RESOURCE_FIELD_MAP_KEY2);
            if (data != null) ParseFieldMap(data);
        }

        /// <summary>
        /// Create an assignment field map from Props data.
        /// </summary>
        public void CreateAssignmentFieldMap(Props props)
        {
            byte[] data = props.GetByteArray(ASSIGNMENT_FIELD_MAP_KEY);
            if (data == null) data = props.GetByteArray(ASSIGNMENT_FIELD_MAP_KEY2);
            if (data != null) ParseFieldMap(data);
        }

        /// <summary>
        /// Parse the raw field map data (28-byte entries).
        /// </summary>
        private void ParseFieldMap(byte[] data)
        {
            int index = 0;
            int lastDataBlockOffset = 0;
            int dataBlockIndex = 0;

            while (index + 28 <= data.Length)
            {
                int dataBlockOffset = ByteArrayHelper.GetShort(data, index + 4);
                int typeValue = ByteArrayHelper.GetInt(data, index + 12);
                int category = ByteArrayHelper.GetShort(data, index + 20);

                int fieldIndex = typeValue & 0x0000FFFF;

                int varDataKey;
                if (m_useTypeAsVarDataKey)
                {
                    varDataKey = typeValue;
                }
                else
                {
                    varDataKey = MppUtility.GetByte(data, index + 6);
                }

                FieldLocation location;

                switch (category)
                {
                    case 0x0B:
                        location = FieldLocation.MetaData;
                        break;
                    case 0x64:
                        location = FieldLocation.MetaData;
                        break;
                    default:
                        if (dataBlockOffset != 65535)
                        {
                            location = FieldLocation.FixedData;
                            if (dataBlockOffset < lastDataBlockOffset)
                            {
                                dataBlockIndex++;
                            }
                            lastDataBlockOffset = dataBlockOffset;

                            int typeSize = GetFieldDataSize(category);
                            if (dataBlockIndex < m_maxFixedDataSize.Length &&
                                dataBlockOffset + typeSize > m_maxFixedDataSize[dataBlockIndex])
                            {
                                m_maxFixedDataSize[dataBlockIndex] = dataBlockOffset + typeSize;
                            }
                        }
                        else
                        {
                            location = varDataKey != 0 ? FieldLocation.VarData : FieldLocation.Unknown;
                        }
                        break;
                }

                var item = new FieldItem
                {
                    FieldIndex = fieldIndex,
                    Location = location,
                    DataBlockIndex = dataBlockIndex,
                    DataBlockOffset = dataBlockOffset,
                    VarDataKey = varDataKey,
                    Category = category
                };

                m_map[fieldIndex] = item;
                index += 28;
            }
        }

        /// <summary>
        /// Estimate the size of a fixed data field based on its category.
        /// </summary>
        private int GetFieldDataSize(int category)
        {
            switch (category)
            {
                case 0x02: return 2;   // short
                case 0x03: return 4;   // int/duration
                case 0x05: return 8;   // rate/numeric (double)
                case 0x08: return 0;   // string (var data)
                case 0x0B: return 0;   // boolean (meta)
                case 0x13: return 4;   // date (timestamp)
                case 0x48: return 16;  // GUID
                case 0x64: return 0;   // boolean (meta)
                case 0x65: return 8;   // work/currency (double)
                case 0x66: return 8;   // units (double)
                case 0x1D: return 4;   // binary
                default: return 4;
            }
        }

        /// <summary>
        /// Get the field item for a given field index.
        /// </summary>
        public FieldItem GetFieldItem(int fieldIndex)
        {
            m_map.TryGetValue(fieldIndex, out FieldItem item);
            return item;
        }

        /// <summary>
        /// Get the fixed data offset for a field, or -1 if not found.
        /// </summary>
        public int GetFixedDataOffset(int fieldIndex)
        {
            if (m_map.TryGetValue(fieldIndex, out FieldItem item) && item.Location == FieldLocation.FixedData)
                return item.DataBlockOffset;
            return -1;
        }

        /// <summary>
        /// Get the fixed data offset for a field in a specific data block, or -1 if not found.
        /// </summary>
        public int GetFixedDataOffset(int fieldIndex, int dataBlockIndex)
        {
            if (m_map.TryGetValue(fieldIndex, out FieldItem item) &&
                item.Location == FieldLocation.FixedData &&
                item.DataBlockIndex == dataBlockIndex)
                return item.DataBlockOffset;
            return -1;
        }

        /// <summary>
        /// Get the var data key for a field, or -1 if not found.
        /// </summary>
        public int GetVarDataKey(int fieldIndex)
        {
            if (m_map.TryGetValue(fieldIndex, out FieldItem item) && item.Location == FieldLocation.VarData)
                return item.VarDataKey;
            return -1;
        }

        /// <summary>
        /// Get the maximum fixed data size for a specific block index.
        /// </summary>
        public int GetMaxFixedDataSize(int blockIndex)
        {
            if (blockIndex >= 0 && blockIndex < m_maxFixedDataSize.Length)
                return m_maxFixedDataSize[blockIndex];
            return 0;
        }

        /// <summary>
        /// Check if we have any field mappings.
        /// </summary>
        public bool HasMappings => m_map.Count > 0;

        /// <summary>
        /// Get all field items.
        /// </summary>
        public IEnumerable<KeyValuePair<int, FieldItem>> Items => m_map;
    }
}
