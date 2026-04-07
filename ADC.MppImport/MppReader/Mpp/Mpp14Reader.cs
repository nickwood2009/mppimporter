using System;
using System.Collections.Generic;
using System.Linq;
using ADC.MppImport.MppReader.Common;
using ADC.MppImport.MppReader.Model;
using ADC.MppImport.Ole2;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Reads MPP14 files (MS Project 2010, 2013, 2016, 2019, 2021+).
    /// Uses FieldMap parsed from Props data to determine actual field offsets.
    /// Ported from org.mpxj.mpp.MPP14Reader
    /// </summary>
    internal class Mpp14Reader : IMppVariantReader
    {
        private const int NULL_TASK_BLOCK_SIZE = 16;

        private MppFileReader m_reader;
        private ProjectFile m_file;
        private CFStorage m_root;
        private CFStorage m_projectDir;
        private Props m_projectProps;
        private Props m_rootProps;
        private DocumentInputStreamFactory m_inputStreamFactory;
        private FieldMap m_taskFieldMap;
        private FieldMap m_resourceFieldMap;
        private FieldMap m_assignmentFieldMap;

        public void Process(MppFileReader reader, ProjectFile file, CompoundFile cf, CFStorage root)
        {
            try
            {
                m_reader = reader;
                m_file = file;
                m_root = root;

                PopulateMemberData();
                ProcessProjectProperties();

                if (!reader.ReadPropertiesOnly)
                {
                    ProcessCalendarData();
                    ProcessResourceData();
                    ProcessTaskData();
                    ProcessAssignmentData();
                    ProcessRelationData();
                }
            }
            catch (Exception ex)
            {
                throw new MppReaderException("Error reading MPP14 file: " + ex.Message, ex);
            }
        }

        private void PopulateMemberData()
        {
            byte[] propsData = MppFileReader.GetStreamData(m_root, "Props14");
            if (propsData == null)
                throw new MppReaderException("Cannot find Props14 stream");

            var props = new Props14(propsData);
            m_rootProps = props;
            m_file.ProjectProperties.ProjectFilePath = props.GetUnicodeString(Props.PROJECT_FILE_PATH);
            m_inputStreamFactory = new DocumentInputStreamFactory(props);

            byte passwordFlag = props.GetByte(Props.PASSWORD_FLAG);
            bool passwordRequiredToRead = (passwordFlag & 0x1) != 0;
            bool encryptionXmlPresent = props.GetByteArray(Props.PROTECTION_PASSWORD_HASH) != null;

            if (passwordRequiredToRead && encryptionXmlPresent && m_reader.RespectPasswordProtection)
                throw new MppReaderException("File is password protected");

            m_projectDir = MppFileReader.GetStorage(m_root, "   114");
            if (m_projectDir == null)
                throw new MppReaderException("Cannot find project directory '   114'");

            byte[] projectPropsData = GetStreamDataWithDecryption(m_projectDir, "Props");
            m_projectProps = projectPropsData != null
                ? new Props14(projectPropsData)
                : new Props14(new byte[0]);

            m_file.ProjectProperties.MppFileType = 14;
            m_file.ProjectProperties.AutoFilter = props.GetBoolean(Props.AUTO_FILTER);

            // Build field maps from Props data (MPP14 uses type as var data key)
            m_taskFieldMap = new FieldMap(true);
            m_taskFieldMap.CreateTaskFieldMap(m_projectProps);

            m_resourceFieldMap = new FieldMap(true);
            m_resourceFieldMap.CreateResourceFieldMap(m_projectProps);

            m_assignmentFieldMap = new FieldMap(true);
            m_assignmentFieldMap.CreateAssignmentFieldMap(m_projectProps);
        }

        private void ProcessProjectProperties()
        {
            ProjectPropertiesReader.Process(m_file, m_projectProps);
        }

        #region Calendar

        private void ProcessCalendarData()
        {
            try
            {
                CFStorage calDir = MppFileReader.GetStorage(m_projectDir, "TBkndCal");
                if (calDir == null) return;

                byte[] fixedMetaData = MppFileReader.GetStreamData(calDir, "FixedMeta");
                byte[] fixedDataBuf = GetStreamDataWithDecryption(calDir, "FixedData");
                byte[] varMetaData = MppFileReader.GetStreamData(calDir, "VarMeta");
                byte[] var2DataBuf = MppFileReader.GetStreamData(calDir, "Var2Data");

                if (fixedMetaData == null || fixedDataBuf == null) return;

                var calFixedMeta = new FixedMeta(fixedMetaData, 10);
                var calFixedData = new FixedData(calFixedMeta, fixedDataBuf);

                IVarMeta calVarMeta = null;
                Var2Data calVarData = null;
                if (varMetaData != null && var2DataBuf != null)
                {
                    calVarMeta = new VarMeta12(varMetaData);
                    calVarData = new Var2Data(calVarMeta, var2DataBuf);
                }

                int items = calFixedMeta.AdjustedItemCount;
                for (int loop = 0; loop < items; loop++)
                {
                    byte[] fixedData = calFixedData.GetByteArrayValue(loop);
                    if (fixedData == null || fixedData.Length < 8) continue;

                    int calUniqueID = ByteArrayHelper.GetInt(fixedData, 0);
                    if (calUniqueID < 1) continue;

                    var calendar = new ProjectCalendar { UniqueID = calUniqueID };

                    if (calVarData != null)
                        calendar.Name = calVarData.GetUnicodeString(calUniqueID, 1);

                    if (fixedData.Length >= 12)
                    {
                        int parentID = ByteArrayHelper.GetInt(fixedData, 8);
                        if (parentID > 0)
                            calendar.ParentCalendarUniqueID = parentID;
                    }

                    ReadCalendarHours(calendar, fixedData);

                    if (calVarData != null)
                        ReadCalendarExceptions(calendar, calVarData, calUniqueID);

                    m_file.Calendars.Add(calendar);
                }
            }
            catch (Exception ex)
            {
                m_file.AddIgnoredError(ex);
            }
        }

        private void ReadCalendarHours(ProjectCalendar calendar, byte[] data)
        {
            int offset = 4;
            if (data.Length < offset + (7 * 60)) return;

            for (int day = 0; day < 7; day++)
            {
                int dayOffset = offset + (day * 60);
                if (dayOffset + 2 > data.Length) break;

                int flags = ByteArrayHelper.GetShort(data, dayOffset);
                var calDay = calendar.Days[day];

                if ((flags & 0x01) != 0)
                {
                    calDay.Type = DayType.Working;
                    for (int range = 0; range < 5; range++)
                    {
                        int rangeOffset = dayOffset + 2 + (range * 8);
                        if (rangeOffset + 8 > data.Length) break;

                        long startMinutes = ByteArrayHelper.GetInt(data, rangeOffset);
                        long endMinutes = ByteArrayHelper.GetInt(data, rangeOffset + 4);

                        if (startMinutes == 0 && endMinutes == 0) break;

                        calDay.Hours.Add(new CalendarHours
                        {
                            Start = TimeSpan.FromMinutes(startMinutes / 60000.0),
                            End = TimeSpan.FromMinutes(endMinutes / 60000.0)
                        });
                    }
                }
                else
                {
                    calDay.Type = (flags & 0x02) != 0 ? DayType.NonWorking : DayType.Default;
                }
            }
        }

        private void ReadCalendarExceptions(ProjectCalendar calendar, Var2Data varData, int calUniqueID)
        {
            byte[] exceptionData = varData.GetByteArray(calUniqueID, 8);
            if (exceptionData == null || exceptionData.Length < 4) return;

            int exCount = ByteArrayHelper.GetInt(exceptionData, 0);
            int offset = 4;

            for (int i = 0; i < exCount && offset + 8 <= exceptionData.Length; i++)
            {
                try
                {
                    var fromDate = MppUtility.GetTimestamp(exceptionData, offset);
                    var toDate = MppUtility.GetTimestamp(exceptionData, offset + 4);

                    if (fromDate.HasValue && toDate.HasValue)
                    {
                        bool working = false;
                        if (offset + 12 <= exceptionData.Length)
                            working = ByteArrayHelper.GetInt(exceptionData, offset + 8) != 0;

                        calendar.Exceptions.Add(new CalendarException
                        {
                            Start = fromDate.Value,
                            End = toDate.Value,
                            Working = working
                        });
                    }
                }
                catch { }
                offset += 12;
            }
        }

        #endregion

        #region Resources

        private void ProcessResourceData()
        {
            try
            {
                CFStorage rscDir = MppFileReader.GetStorage(m_projectDir, "TBkndRsc");
                if (rscDir == null) return;

                byte[] varMetaData = MppFileReader.GetStreamData(rscDir, "VarMeta");
                byte[] var2DataBuf = MppFileReader.GetStreamData(rscDir, "Var2Data");
                byte[] fixedMetaData = MppFileReader.GetStreamData(rscDir, "FixedMeta");
                byte[] fixedDataBuf = GetStreamDataWithDecryption(rscDir, "FixedData");

                if (fixedMetaData == null || fixedDataBuf == null) return;

                var rscFixedMeta = new FixedMeta(fixedMetaData, 37);
                var rscFixedData = new FixedData(rscFixedMeta, fixedDataBuf);

                IVarMeta rscVarMeta = null;
                Var2Data rscVarData = null;
                if (varMetaData != null && var2DataBuf != null)
                {
                    rscVarMeta = new VarMeta12(varMetaData);
                    rscVarData = new Var2Data(rscVarMeta, var2DataBuf);
                }

                int itemCount = rscFixedMeta.AdjustedItemCount;
                var fm = m_resourceFieldMap;

                for (int loop = 0; loop < itemCount; loop++)
                {
                    byte[] data = rscFixedData.GetByteArrayValue(loop);
                    if (data == null || data.Length < 8) continue;

                    int uniqueID = ReadFixedShort(data, fm, (int)ResourceFieldIndex.UniqueID, 0);
                    if (uniqueID < 1) continue;

                    var resource = new Resource();
                    resource.UniqueID = uniqueID;
                    resource.ID = ReadFixedInt(data, fm, (int)ResourceFieldIndex.ID, 4);

                    // Var data fields
                    if (rscVarData != null)
                    {
                        resource.Name = ReadVarString(rscVarData, uniqueID, fm, (int)ResourceFieldIndex.Name);
                        resource.Initials = ReadVarString(rscVarData, uniqueID, fm, (int)ResourceFieldIndex.Initials);
                        resource.Group = ReadVarString(rscVarData, uniqueID, fm, (int)ResourceFieldIndex.Group);
                        resource.Code = ReadVarString(rscVarData, uniqueID, fm, (int)ResourceFieldIndex.Code);
                        resource.EmailAddress = ReadVarString(rscVarData, uniqueID, fm, (int)ResourceFieldIndex.EmailAddress);
                        resource.Notes = ReadVarString(rscVarData, uniqueID, fm, (int)ResourceFieldIndex.Notes);
                    }

                    // Fixed data fields
                    resource.MaxUnits = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.MaxUnits) / 100.0;
                    resource.PercentWorkComplete = ReadFixedShort(data, fm, (int)ResourceFieldIndex.PercentWorkComplete);

                    int accrueVal = ReadFixedShort(data, fm, (int)ResourceFieldIndex.AccrueAt);
                    if (accrueVal >= 1 && accrueVal <= 3)
                        resource.AccrueAt = (AccrueType)accrueVal;

                    double stdRate = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.StandardRate);
                    if (stdRate != 0) resource.StandardRate = new Rate(stdRate / 100.0, TimeUnit.Hours);

                    double otRate = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.OvertimeRate);
                    if (otRate != 0) resource.OvertimeRate = new Rate(otRate / 100.0, TimeUnit.Hours);

                    resource.CreationDate = ReadFixedTimestamp(data, fm, (int)ResourceFieldIndex.CreationDate);

                    int calID = ReadFixedInt(data, fm, (int)ResourceFieldIndex.CalendarUniqueID);
                    if (calID > 0) resource.CalendarUniqueID = calID;

                    double work = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.Work);
                    if (work != 0) resource.Work = MppUtility.GetWorkDuration(work);

                    double actualWork = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.ActualWork);
                    if (actualWork != 0) resource.ActualWork = MppUtility.GetWorkDuration(actualWork);

                    double remainWork = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.RemainingWork);
                    if (remainWork != 0) resource.RemainingWork = MppUtility.GetWorkDuration(remainWork);

                    double cost = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.Cost);
                    if (cost != 0) resource.Cost = cost / 100.0;

                    double actualCost = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.ActualCost);
                    if (actualCost != 0) resource.ActualCost = actualCost / 100.0;

                    double remainCost = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.RemainingCost);
                    if (remainCost != 0) resource.RemainingCost = remainCost / 100.0;

                    double costPerUse = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.CostPerUse);
                    if (costPerUse != 0) resource.CostPerUse = costPerUse / 100.0;

                    resource.Active = true;
                    m_file.Resources.Add(resource);
                }
            }
            catch (Exception ex)
            {
                m_file.AddIgnoredError(ex);
            }
        }

        #endregion

        #region Tasks

        private void ProcessTaskData()
        {
            try
            {
                CFStorage taskDir = MppFileReader.GetStorage(m_projectDir, "TBkndTask");
                if (taskDir == null) return;

                byte[] varMetaData = MppFileReader.GetStreamData(taskDir, "VarMeta");
                byte[] var2DataBuf = MppFileReader.GetStreamData(taskDir, "Var2Data");
                byte[] fixedMetaData = MppFileReader.GetStreamData(taskDir, "FixedMeta");
                byte[] fixedDataBuf = MppFileReader.GetStreamData(taskDir, "FixedData");
                byte[] fixed2DataBuf = MppFileReader.GetStreamData(taskDir, "Fixed2Data");
                if (fixed2DataBuf == null)
                    fixed2DataBuf = MppFileReader.GetStreamData(taskDir, "FixedData2");

                if (fixedMetaData == null || fixedDataBuf == null) return;

                var taskFixedMeta = new FixedMeta(fixedMetaData, 47);
                int maxSize = m_taskFieldMap.GetMaxFixedDataSize(0);
                if (maxSize < 4) maxSize = 200;
                var taskFixedData = new FixedData(taskFixedMeta, fixedDataBuf, maxSize);

                var properties = m_file.ProjectProperties;
                var fm = m_taskFieldMap;

                // Second fixed data block (contains fields like DurationUnits, OutlineLevel)
                FixedData taskFixed2Data = null;
                if (fixed2DataBuf != null)
                {
                    int maxSize2 = m_taskFieldMap.GetMaxFixedDataSize(1);
                    if (maxSize2 < 4) maxSize2 = 768;
                    taskFixed2Data = new FixedData(taskFixedMeta, fixed2DataBuf, maxSize2);
                    m_file.DiagnosticMessages.Add(string.Format("Fixed2Data: found, {0} bytes, maxSize2={1}", fixed2DataBuf.Length, maxSize2));
                }
                else
                {
                    m_file.DiagnosticMessages.Add("Fixed2Data: NOT FOUND");
                }

                IVarMeta taskVarMeta = null;
                Var2Data taskVarData = null;
                if (varMetaData != null && var2DataBuf != null)
                {
                    taskVarMeta = new VarMeta12(varMetaData);
                    taskVarData = new Var2Data(taskVarMeta, var2DataBuf);
                }

                int itemCount = taskFixedMeta.AdjustedItemCount;

                // Build lookup table (GUID → text) from TBkndOutlCode for resolving
                // custom field values that store GUID references instead of direct text
                var lookupGuidToText = new Dictionary<string, string>();
                CFStorage outlCodeDir = MppFileReader.GetStorage(m_projectDir, "TBkndOutlCode");
                if (outlCodeDir != null)
                {
                    byte[] ocVarMetaBuf = MppFileReader.GetStreamData(outlCodeDir, "VarMeta");
                    byte[] ocVar2DataBuf = MppFileReader.GetStreamData(outlCodeDir, "Var2Data");
                    byte[] ocFixed2Buf = MppFileReader.GetStreamData(outlCodeDir, "Fixed2Data");
                    if (ocVarMetaBuf != null && ocVar2DataBuf != null && ocFixed2Buf != null)
                    {
                        var ocVarMeta = new VarMeta12(ocVarMetaBuf);
                        var ocVarData = new Var2Data(ocVarMeta, ocVar2DataBuf);
                        // Fixed2Data: 34 bytes per entry (16-byte entry GUID + 16-byte parent GUID + 2-byte flags)
                        // First 3 entries are field headers; value entries start at index 3
                        // Value entry at index E corresponds to VarData UID = E - 3
                        const int ENTRY_SIZE = 34;
                        int totalEntries = ocFixed2Buf.Length / ENTRY_SIZE;
                        int valueStartIdx = 3; // skip field header entries
                        for (int e = valueStartIdx; e < totalEntries; e++)
                        {
                            int entryOffset = e * ENTRY_SIZE;
                            if (entryOffset + 16 > ocFixed2Buf.Length) break;
                            // Extract 16-byte GUID from the entry
                            string guidKey = BitConverter.ToString(ocFixed2Buf, entryOffset, 16);
                            // The corresponding VarData UID
                            int ocUid = e - valueStartIdx;
                            // Read text value from VarData (type 0x16400016)
                            string text = ocVarData.GetUnicodeString(ocUid, 0x16400016);
                            if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(guidKey))
                            {
                                lookupGuidToText[guidKey] = text;
                            }
                        }
                        m_file.DiagnosticMessages.Add(string.Format(
                            "  [OUTLCODE] Built lookup table: {0} GUID→text entries from {1} Fixed2Data entries",
                            lookupGuidToText.Count, totalEntries));
                    }
                }

                // Diagnostic: Active field map info
                var activeFieldItem = fm.GetFieldItem((int)TaskFieldIndex.Active);
                if (activeFieldItem != null)
                    m_file.DiagnosticMessages.Add(string.Format("Active field: Location={0}, Category=0x{1:X2}, MetaIdx={2}",
                        activeFieldItem.Location, activeFieldItem.Category, activeFieldItem.MetaDataIndex));
                else
                    m_file.DiagnosticMessages.Add("Active field: NOT in field map");

                var knownIndices = new HashSet<int>(
                    Enum.GetValues(typeof(TaskFieldIndex)).Cast<int>());
                m_file.DiagnosticMessages.Add(string.Format("Task field map: {0} total entries", fm.Items.Count()));

                // Parse field name aliases (user-defined custom column names)
                // MPP14 stores aliases in TBkndTask/Props stream, not in the project-level Props
                var aliases = new Dictionary<int, string>();
                byte[] taskPropsData = GetStreamDataWithDecryption(taskDir, "Props");
                if (taskPropsData != null && taskPropsData.Length > 0)
                {
                    var taskProps = new Props14(taskPropsData);
                    aliases = ReadTaskFieldAliases(taskProps);

                    // If standard key not found, try CUSTOM_FIELDS key
                    if (aliases.Count == 0)
                    {
                        byte[] cfData = taskProps.GetByteArray(Props.CUSTOM_FIELDS);
                        if (cfData != null && cfData.Length > 0)
                            aliases = TryParseAliasBlock(cfData);
                    }
                }
                // Fallback: try project props and root props
                if (aliases.Count == 0)
                    aliases = ReadTaskFieldAliases(m_projectProps);
                if (aliases.Count == 0 && m_rootProps != null)
                    aliases = ReadTaskFieldAliases(m_rootProps);
                if (aliases.Count > 0)
                {
                    m_file.DiagnosticMessages.Add(string.Format("Task field aliases: {0} entries", aliases.Count));
                    foreach (var a in aliases)
                    {
                        string fieldLabel = knownIndices.Contains(a.Key) ? ((TaskFieldIndex)a.Key).ToString() : "CUSTOM";
                        m_file.DiagnosticMessages.Add(string.Format("  FieldIdx={0} ({1}) => \"{2}\"", a.Key, fieldLabel, a.Value));
                    }
                }
                else
                {
                    m_file.DiagnosticMessages.Add("Task field aliases: NONE");
                }


                // Identify custom fields to read (VarData string, int, double, date types)
                var customFields = new List<KeyValuePair<int, FieldItem>>();
                foreach (var kvp in fm.Items)
                {
                    if (knownIndices.Contains(kvp.Key)) continue;
                    if (kvp.Value.Location == FieldLocation.VarData &&
                        (kvp.Value.Category == 0x08 || kvp.Value.Category == 0x03 ||
                         kvp.Value.Category == 0x05 || kvp.Value.Category == 0x65 ||
                         kvp.Value.Category == 0x13))
                    {
                        customFields.Add(kvp);
                    }
                }

                // Build task map: skip first 3 items (not real tasks)
                var taskMap = new SortedDictionary<int, int>();
                for (int loop = itemCount - 1; loop > 2; loop--)
                {
                    byte[] data = taskFixedData.GetByteArrayValue(loop);
                    if (data == null) continue;

                    byte[] metaData = taskFixedMeta.GetByteArrayValue(loop);
                    if (metaData == null) continue;

                    int flags = ByteArrayHelper.GetInt(metaData, 0);
                    if ((flags & 0x02) != 0) continue; // deleted

                    int uniqueID;
                    if (data.Length == NULL_TASK_BLOCK_SIZE)
                        uniqueID = ByteArrayHelper.GetInt(data, 0);
                    else
                        uniqueID = ReadFixedInt(data, fm, (int)TaskFieldIndex.UniqueID, 0);

                    if (!taskMap.ContainsKey(uniqueID))
                        taskMap[uniqueID] = loop;
                }

                foreach (var entry in taskMap)
                {
                    int uniqueID = entry.Key;
                    int index = entry.Value;

                    byte[] data = taskFixedData.GetByteArrayValue(index);
                    if (data == null) continue;
                    byte[] data2 = taskFixed2Data != null ? taskFixed2Data.GetByteArrayValue(index) : null;

                    // Null/placeholder tasks
                    if (data.Length == NULL_TASK_BLOCK_SIZE)
                    {
                        var nullTask = new Task();
                        nullTask.UniqueID = ByteArrayHelper.GetInt(data, 0);
                        nullTask.ID = ByteArrayHelper.GetInt(data, 4);
                        m_file.Tasks.Add(nullTask);
                        continue;
                    }

                    if (data.Length < 8) continue;

                    var task = new Task();
                    task.UniqueID = uniqueID;
                    task.ID = ReadFixedInt(data, fm, (int)TaskFieldIndex.ID, 4);

                    // Dates from fixed data
                    task.Start = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.Start);
                    task.Finish = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.Finish);
                    task.ActualStart = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.ActualStart);
                    task.ActualFinish = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.ActualFinish);
                    task.EarlyStart = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.EarlyStart);
                    task.EarlyFinish = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.EarlyFinish);
                    task.LateStart = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.LateStart);
                    task.LateFinish = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.LateFinish);
                    task.ConstraintDate = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.ConstraintDate);
                    task.CreateDate = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.Created);
                    task.Deadline = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.Deadline);

                    // Hierarchy
                    int parentUniqueID = ReadFixedInt(data, fm, (int)TaskFieldIndex.ParentTaskUniqueID);
                    if (parentUniqueID >= 0 && parentUniqueID != uniqueID)
                        task.ParentTaskUniqueID = parentUniqueID;
                    task.OutlineLevel = ReadFixedShortFromBlock(data, data2, fm, (int)TaskFieldIndex.OutlineLevel);

                    // Duration units: try VarData/FixedData first, then infer from date range
                    int durationUnitsValue = ReadFieldShort(data, data2, taskVarData, uniqueID, fm, (int)TaskFieldIndex.DurationUnits);
                    var durationUnits = MppUtility.GetDurationTimeUnits(durationUnitsValue, properties.DefaultDurationUnits);

                    int rawDuration = ReadFixedInt(data, fm, (int)TaskFieldIndex.Duration);

                    // If DurationUnits was not found (defaulted to Days), infer from date range.
                    // eDays raw = N * 1440 * 10 = N * 14400; Days raw = N * 480 * 10 = N * 4800.
                    // If rawDuration / 14400 ≈ calendar day span → ElapsedDays.
                    if (durationUnitsValue == 0 && rawDuration > 0 && task.Start.HasValue && task.Finish.HasValue)
                    {
                        double calendarDays = (task.Finish.Value.Date - task.Start.Value.Date).TotalDays;
                        double eDaysValue = rawDuration / 14400.0;
                        double roundedEDays = Math.Round(eDaysValue);
                        // eDays in MPP are whole numbers, so eDaysValue must be near-integer
                        if (calendarDays > 0 && Math.Abs(eDaysValue - roundedEDays) < 0.01
                            && Math.Abs(roundedEDays - calendarDays) < 0.5)
                        {
                            durationUnits = TimeUnit.ElapsedDays;
                        }
                    }
                    task.Duration = MppUtility.GetAdjustedDuration(properties, rawDuration, durationUnits);

                    int rawActualDuration = ReadFixedInt(data, fm, (int)TaskFieldIndex.ActualDuration);
                    task.ActualDuration = MppUtility.GetAdjustedDuration(properties, rawActualDuration, durationUnits);

                    int rawRemainingDuration = ReadFixedInt(data, fm, (int)TaskFieldIndex.RemainingDuration);
                    task.RemainingDuration = MppUtility.GetAdjustedDuration(properties, rawRemainingDuration, durationUnits);

                    int rawFreeSlack = ReadFixedInt(data, fm, (int)TaskFieldIndex.FreeSlack);
                    task.FreeSlack = MppUtility.GetAdjustedDuration(properties, rawFreeSlack, durationUnits);

                    // Constraint
                    int constraintType = ReadFixedShort(data, fm, (int)TaskFieldIndex.ConstraintType);
                    if (constraintType >= 0 && constraintType <= 7)
                        task.ConstraintType = (ConstraintType)constraintType;

                    // Priority, Percent Complete, Type
                    task.Priority = ReadFixedShort(data, fm, (int)TaskFieldIndex.Priority);
                    task.PercentComplete = ReadFixedShort(data, fm, (int)TaskFieldIndex.PercentComplete);
                    task.PercentWorkComplete = ReadFixedShort(data, fm, (int)TaskFieldIndex.PercentWorkComplete);

                    int taskType = ReadFixedShort(data, fm, (int)TaskFieldIndex.Type);
                    if (taskType >= 0 && taskType <= 2) task.Type = (TaskType)taskType;

                    // Calendar
                    int calID = ReadFixedInt(data, fm, (int)TaskFieldIndex.CalendarUniqueID);
                    if (calID > 0) task.CalendarUniqueID = calID;

                    // Work fields (WORK data type: double in ms, / 60000 = hours)
                    double work = ReadFixedDouble(data, fm, (int)TaskFieldIndex.Work);
                    if (work != 0) task.Work = MppUtility.GetWorkDuration(work);

                    double actualWork = ReadFixedDouble(data, fm, (int)TaskFieldIndex.ActualWork);
                    if (actualWork != 0) task.ActualWork = MppUtility.GetWorkDuration(actualWork);

                    double remainWork = ReadFixedDouble(data, fm, (int)TaskFieldIndex.RemainingWork);
                    if (remainWork != 0) task.RemainingWork = MppUtility.GetWorkDuration(remainWork);

                    // Cost fields (stored as double, 1/100 units)
                    double cost = ReadFixedDouble(data, fm, (int)TaskFieldIndex.Cost);
                    if (cost != 0) task.Cost = cost / 100.0;

                    double fixedCost = ReadFixedDouble(data, fm, (int)TaskFieldIndex.FixedCost);
                    if (fixedCost != 0) task.FixedCost = fixedCost / 100.0;

                    double actualCost = ReadFixedDouble(data, fm, (int)TaskFieldIndex.ActualCost);
                    if (actualCost != 0) task.ActualCost = actualCost / 100.0;

                    double remainCost = ReadFixedDouble(data, fm, (int)TaskFieldIndex.RemainingCost);
                    if (remainCost != 0) task.RemainingCost = remainCost / 100.0;

                    // Var data fields (name, WBS, notes)
                    if (taskVarData != null)
                    {
                        task.Name = ReadVarString(taskVarData, uniqueID, fm, (int)TaskFieldIndex.Name);
                        task.WBS = ReadVarString(taskVarData, uniqueID, fm, (int)TaskFieldIndex.WBS);
                        task.Notes = ReadVarString(taskVarData, uniqueID, fm, (int)TaskFieldIndex.Notes);
                    }

                    // Read custom fields from VarData
                    if (taskVarData != null && customFields.Count > 0)
                    {
                        foreach (var cf in customFields)
                        {
                            int cfIdx = cf.Key;
                            int varKey = cf.Value.VarDataKey;
                            string cfName;
                            if (!aliases.TryGetValue(cfIdx, out cfName))
                                cfName = GetStandardCustomFieldName(cfIdx);

                            object value = null;
                            switch (cf.Value.Category)
                            {
                                case 0x08: // string
                                    // First get raw bytes to detect lookup table GUID references
                                    byte[] rawCf = taskVarData.GetByteArray(uniqueID, varKey);
                                    if (rawCf != null && rawCf.Length >= 22 && rawCf[0] == 0x01)
                                    {
                                        // Lookup table reference: [01][07][4-byte][16-byte GUID][...]
                                        // Extract 16-byte GUID from bytes 6-21
                                        string guidKey = BitConverter.ToString(rawCf, 6, 16);
                                        string resolved;
                                        if (lookupGuidToText.TryGetValue(guidKey, out resolved))
                                        {
                                            value = resolved;
                                        }
                                        else
                                        {
                                            m_file.DiagnosticMessages.Add(string.Format(
                                                "  [CF-LUT-MISS] UID={0} \"{1}\" guidKey={2} (not in lookup table)",
                                                uniqueID, cfName, guidKey));
                                        }
                                    }
                                    else
                                    {
                                        // Normal text value
                                        string s = rawCf != null ? MppUtility.GetUnicodeString(rawCf, 0) : null;
                                        if (!string.IsNullOrEmpty(s))
                                            value = s;
                                    }
                                    break;
                                case 0x03: // int/duration
                                    int iv = taskVarData.GetInt(uniqueID, varKey);
                                    if (iv != 0) value = iv;
                                    break;
                                case 0x05: // double/numeric
                                case 0x65: // work/currency
                                    double dv = taskVarData.GetDouble(uniqueID, varKey);
                                    if (dv != 0) value = dv;
                                    break;
                                case 0x13: // date
                                    DateTime? dt = taskVarData.GetTimestamp(uniqueID, varKey);
                                    if (dt.HasValue) value = dt.Value;
                                    break;
                            }

                            if (value != null)
                                task.CustomFields[cfName] = value;
                        }
                    }

                    // Milestone: a task is a milestone if its duration is zero
                    task.Milestone = (rawDuration == 0);

                    // Read Active field from metadata booleans.
                    // Each boolean field uses 2 bits (value + defined), starting at byte 8.
                    bool isActive = true;
                    var activeItem = fm.GetFieldItem((int)TaskFieldIndex.Active);
                    if (activeItem != null && activeItem.Location == FieldLocation.MetaData && activeItem.MetaDataIndex >= 0)
                    {
                        byte[] taskMeta = taskFixedMeta.GetByteArrayValue(index);
                        if (taskMeta != null)
                        {
                            int totalBit = activeItem.MetaDataIndex * 2;
                            int byteIdx = 8 + (totalBit / 8);
                            int bitMask = 1 << (totalBit % 8);
                            if (byteIdx < taskMeta.Length)
                                isActive = (taskMeta[byteIdx] & bitMask) != 0;
                        }
                    }
                    task.Active = isActive;

                    if (task.Name == null && task.Start == null && task.Finish == null)
                        continue;

                    m_file.Tasks.Add(task);
                }
            }
            catch (Exception ex)
            {
                m_file.AddIgnoredError(ex);
            }
        }

        #endregion

        #region Assignments

        private void ProcessAssignmentData()
        {
            try
            {
                CFStorage assnDir = MppFileReader.GetStorage(m_projectDir, "TBkndAssn");
                if (assnDir == null) return;

                byte[] fixedMetaData = MppFileReader.GetStreamData(assnDir, "FixedMeta");
                byte[] fixedDataBuf = GetStreamDataWithDecryption(assnDir, "FixedData");
                if (fixedDataBuf == null) return;

                byte[] varMetaData = MppFileReader.GetStreamData(assnDir, "VarMeta");
                byte[] var2DataBuf = MppFileReader.GetStreamData(assnDir, "Var2Data");

                int maxSize = m_assignmentFieldMap.GetMaxFixedDataSize(0);
                if (maxSize < 20) maxSize = 110;
                var assnFixedData = new FixedData(maxSize, fixedDataBuf);

                IVarMeta assnVarMeta = null;
                Var2Data assnVarData = null;
                if (varMetaData != null && var2DataBuf != null)
                {
                    assnVarMeta = new VarMeta12(varMetaData);
                    assnVarData = new Var2Data(assnVarMeta, var2DataBuf);
                }

                var fm = m_assignmentFieldMap;
                int itemCount = assnFixedData.ItemCount;

                for (int loop = 0; loop < itemCount; loop++)
                {
                    byte[] data = assnFixedData.GetByteArrayValue(loop);
                    if (data == null || data.Length < 12) continue;

                    int uniqueID = ReadFixedInt(data, fm, (int)AssignmentFieldIndex.UniqueID, 0);
                    if (uniqueID < 1) continue;

                    int taskUniqueID = ReadFixedInt(data, fm, (int)AssignmentFieldIndex.TaskUniqueID, 4);
                    int resourceUniqueID = ReadFixedInt(data, fm, (int)AssignmentFieldIndex.ResourceUniqueID, 8);
                    if (taskUniqueID < 1) continue;

                    var assignment = new ResourceAssignment();
                    assignment.UniqueID = uniqueID;
                    assignment.TaskUniqueID = taskUniqueID;
                    assignment.ResourceUniqueID = resourceUniqueID > 0 ? resourceUniqueID : (int?)null;

                    assignment.Start = ReadFixedTimestamp(data, fm, (int)AssignmentFieldIndex.Start, 12);
                    assignment.Finish = ReadFixedTimestamp(data, fm, (int)AssignmentFieldIndex.Finish, 16);

                    double units = ReadFixedDouble(data, fm, (int)AssignmentFieldIndex.Units, 46);
                    if (units != 0) assignment.Units = units / 100.0;

                    double work = ReadFixedDouble(data, fm, (int)AssignmentFieldIndex.Work, 54);
                    if (work != 0) assignment.Work = MppUtility.GetWorkDuration(work);

                    double actualWork = ReadFixedDouble(data, fm, (int)AssignmentFieldIndex.ActualWork, 62);
                    if (actualWork != 0) assignment.ActualWork = MppUtility.GetWorkDuration(actualWork);

                    double remainWork = ReadFixedDouble(data, fm, (int)AssignmentFieldIndex.RemainingWork, 78);
                    if (remainWork != 0) assignment.RemainingWork = MppUtility.GetWorkDuration(remainWork);

                    double cost = ReadFixedDouble(data, fm, (int)AssignmentFieldIndex.Cost, 86);
                    if (cost != 0) assignment.Cost = cost / 100.0;

                    double actualCost = ReadFixedDouble(data, fm, (int)AssignmentFieldIndex.ActualCost, 94);
                    if (actualCost != 0) assignment.ActualCost = actualCost / 100.0;

                    m_file.Assignments.Add(assignment);
                }
            }
            catch (Exception ex)
            {
                m_file.AddIgnoredError(ex);
            }
        }

        #endregion

        #region Relations

        private void ProcessRelationData()
        {
            try
            {
                CFStorage linkDir = MppFileReader.GetStorage(m_projectDir, "TBkndLink");
                bool useCons = false;

                if (linkDir == null)
                {
                    // Newer MPP formats store links in TBkndCons instead of TBkndLink
                    linkDir = MppFileReader.GetStorage(m_projectDir, "TBkndCons");
                    useCons = true;
                    if (linkDir == null) return;
                }

                byte[] linkFixedDataBuf = MppFileReader.GetStreamData(linkDir, "FixedData");
                if (linkFixedDataBuf == null || linkFixedDataBuf.Length < 16) return;

                int relationsFound = 0;

                if (useCons)
                {
                    // TBkndCons format: packed 20-byte records in raw FixedData buffer
                    const int RECORD_SIZE = 20;
                    int recordCount = linkFixedDataBuf.Length / RECORD_SIZE;

                    for (int i = 0; i < recordCount; i++)
                    {
                        int offset = i * RECORD_SIZE;
                        if (offset + 16 > linkFixedDataBuf.Length) break;

                        int sourceTaskUniqueID = ByteArrayHelper.GetInt(linkFixedDataBuf, offset + 4);
                        int targetTaskUniqueID = ByteArrayHelper.GetInt(linkFixedDataBuf, offset + 8);
                        int relationType = ByteArrayHelper.GetShort(linkFixedDataBuf, offset + 12);
                        int lagDuration = (offset + 18 <= linkFixedDataBuf.Length)
                            ? ByteArrayHelper.GetInt(linkFixedDataBuf, offset + 14) : 0;

                        if (sourceTaskUniqueID < 1 || targetTaskUniqueID < 1) continue;
                        if (sourceTaskUniqueID == targetTaskUniqueID) continue;
                        if (relationType < 0 || relationType > 3) continue;

                        AddRelation(sourceTaskUniqueID, targetTaskUniqueID, relationType, lagDuration);
                        relationsFound++;
                    }
                }
                else
                {
                    // TBkndLink format: records split by FixedMeta
                    byte[] linkFixedMetaData = MppFileReader.GetStreamData(linkDir, "FixedMeta");
                    if (linkFixedMetaData == null) return;

                    var linkFixedMeta = new FixedMeta(linkFixedMetaData, 13);
                    var linkFixedData = new FixedData(linkFixedMeta, linkFixedDataBuf);
                    int itemCount = linkFixedMeta.AdjustedItemCount;

                    for (int loop = 0; loop < itemCount; loop++)
                    {
                        byte[] data = linkFixedData.GetByteArrayValue(loop);
                        if (data == null || data.Length < 16) continue;

                        byte[] metaData = linkFixedMeta.GetByteArrayValue(loop);
                        if (metaData == null) continue;

                        int flags = ByteArrayHelper.GetInt(metaData, 0);
                        if ((flags & 0x02) != 0) continue;

                        int sourceTaskUniqueID = ByteArrayHelper.GetInt(data, 4);
                        int targetTaskUniqueID = ByteArrayHelper.GetInt(data, 8);
                        int relationType = ByteArrayHelper.GetShort(data, 12);
                        int lagDuration = data.Length > 16 ? ByteArrayHelper.GetInt(data, 14) : 0;

                        if (sourceTaskUniqueID < 1 || targetTaskUniqueID < 1) continue;
                        if (relationType < 0 || relationType > 3) continue;

                        AddRelation(sourceTaskUniqueID, targetTaskUniqueID, relationType, lagDuration);
                        relationsFound++;
                    }
                }
            }
            catch (Exception ex)
            {
                m_file.AddIgnoredError(ex);
            }
        }

        private void AddRelation(int sourceTaskUniqueID, int targetTaskUniqueID, int relationType, int lagDuration)
        {
            var relation = new Relation();
            relation.SourceTaskUniqueID = sourceTaskUniqueID;
            relation.TargetTaskUniqueID = targetTaskUniqueID;

            if (relationType >= 0 && relationType <= 3)
                relation.Type = (RelationType)relationType;

            if (lagDuration != 0)
            {
                var du = m_file.ProjectProperties.DefaultDurationUnits;
                relation.Lag = MppUtility.GetAdjustedDuration(m_file.ProjectProperties, lagDuration, du);
            }

            var sourceTask = m_file.Tasks.FirstOrDefault(t => t.UniqueID == sourceTaskUniqueID);
            if (sourceTask != null) sourceTask.Successors.Add(relation);

            var targetTask = m_file.Tasks.FirstOrDefault(t => t.UniqueID == targetTaskUniqueID);
            if (targetTask != null) targetTask.Predecessors.Add(relation);
        }

        #endregion

        #region FieldMap Helpers

        private int ReadFixedInt(byte[] data, FieldMap fm, int fieldIndex, int defaultOffset = -1)
        {
            int offset = fm.GetFixedDataOffset(fieldIndex);
            if (offset < 0) offset = defaultOffset;
            if (offset < 0 || offset + 4 > data.Length) return 0;
            return ByteArrayHelper.GetInt(data, offset);
        }

        private int ReadFixedShort(byte[] data, FieldMap fm, int fieldIndex, int defaultOffset = -1)
        {
            int offset = fm.GetFixedDataOffset(fieldIndex);
            if (offset < 0) offset = defaultOffset;
            if (offset < 0 || offset + 2 > data.Length) return 0;
            return ByteArrayHelper.GetShort(data, offset);
        }

        /// <summary>
        /// Reads a short from the correct data block (block 0 or block 1) based on the field map.
        /// Falls back to block 0 if block 1 is not available.
        /// </summary>
        private int ReadFixedShortFromBlock(byte[] data0, byte[] data1, FieldMap fm, int fieldIndex)
        {
            var item = fm.GetFieldItem(fieldIndex);
            if (item != null && item.Location == FieldLocation.FixedData)
            {
                byte[] targetData = (item.DataBlockIndex == 1 && data1 != null) ? data1 : data0;
                int offset = item.DataBlockOffset;
                if (offset >= 0 && offset + 2 <= targetData.Length)
                    return ByteArrayHelper.GetShort(targetData, offset);
            }
            return 0;
        }

        /// <summary>
        /// Reads a short value from VarData, FixedData block 0, or FixedData block 1
        /// depending on where the field map says the field lives.
        /// </summary>
        private int ReadFieldShort(byte[] data0, byte[] data1, Var2Data varData, int uniqueID, FieldMap fm, int fieldIndex)
        {
            var item = fm.GetFieldItem(fieldIndex);
            if (item == null) return 0;

            if (item.Location == FieldLocation.VarData && varData != null)
            {
                return varData.GetShort(uniqueID, item.VarDataKey);
            }

            if (item.Location == FieldLocation.FixedData)
            {
                byte[] targetData = (item.DataBlockIndex == 1 && data1 != null) ? data1 : data0;
                int offset = item.DataBlockOffset;
                if (offset >= 0 && offset + 2 <= targetData.Length)
                    return ByteArrayHelper.GetShort(targetData, offset);
            }

            return 0;
        }

        private double ReadFixedDouble(byte[] data, FieldMap fm, int fieldIndex, int defaultOffset = -1)
        {
            int offset = fm.GetFixedDataOffset(fieldIndex);
            if (offset < 0) offset = defaultOffset;
            if (offset < 0 || offset + 8 > data.Length) return 0;
            return MppUtility.GetDouble(data, offset);
        }

        private DateTime? ReadFixedTimestamp(byte[] data, FieldMap fm, int fieldIndex, int defaultOffset = -1)
        {
            int offset = fm.GetFixedDataOffset(fieldIndex);
            if (offset < 0) offset = defaultOffset;
            if (offset < 0 || offset + 4 > data.Length) return null;
            return MppUtility.GetTimestamp(data, offset);
        }

        private string ReadVarString(Var2Data varData, int uniqueID, FieldMap fm, int fieldIndex)
        {
            int key = fm.GetVarDataKey(fieldIndex);
            if (key < 0) key = fieldIndex; // fall back to field index as key
            return varData.GetUnicodeString(uniqueID, key);
        }

        #endregion

        /// <summary>
        /// Map a field index to a standard MPP custom field name.
        /// MPP custom text fields: Text1=51, Text2=54, Text3=57, ... (stride 3, up to Text30)
        /// MPP custom number fields: Number1=259, Number2=260, ... (stride 1, up to Number20)
        /// MPP custom date fields: Date1=265, Date2=266, ... (stride 1, up to Date10)
        /// MPP custom flag fields: Flag1=299, Flag2=300, ... (stride 1, up to Flag20)
        /// MPP custom cost fields: Cost1=279, Cost2=280, ... (stride 1, up to Cost10)
        /// MPP custom start fields: Start1=52, Start2=55, ... (stride 3, up to Start10)
        /// MPP custom finish fields: Finish1=53, Finish2=56, ... (stride 3, up to Finish10)
        /// MPP custom duration fields: Duration1=244, Duration2=245, ... (stride 1, up to Duration10)
        /// </summary>
        private static string GetStandardCustomFieldName(int fieldIndex)
        {
            // Text1-30: starts at 51, stride 3
            if (fieldIndex >= 51 && fieldIndex <= 51 + 29 * 3)
            {
                int offset = fieldIndex - 51;
                if (offset % 3 == 0) return string.Format("Text{0}", offset / 3 + 1);
                if (offset % 3 == 1) return string.Format("Start{0}", offset / 3 + 1);
                if (offset % 3 == 2) return string.Format("Finish{0}", offset / 3 + 1);
            }
            // Duration1-10: 244-253
            if (fieldIndex >= 244 && fieldIndex <= 253)
                return string.Format("Duration{0}", fieldIndex - 244 + 1);
            // Date1-10: 265-274  (check before Number since ranges overlap)
            if (fieldIndex >= 265 && fieldIndex <= 274)
                return string.Format("Date{0}", fieldIndex - 265 + 1);
            // Number1-20: 259-264, 275-278 (skipping Date range)
            if (fieldIndex >= 259 && fieldIndex <= 278)
                return string.Format("Number{0}", fieldIndex - 259 + 1);
            // Cost1-10: 279-288
            if (fieldIndex >= 279 && fieldIndex <= 288)
                return string.Format("Cost{0}", fieldIndex - 279 + 1);
            // Flag1-20: 299-318
            if (fieldIndex >= 299 && fieldIndex <= 318)
                return string.Format("Flag{0}", fieldIndex - 299 + 1);
            return string.Format("Field_{0}", fieldIndex);
        }

        /// <summary>
        /// Try to parse an alias block from raw data.
        /// Format: [count:4] [entries: count * (fieldTypeId:4, stringOffset:4)] [unicode strings]
        /// Field type IDs follow the 0x0B40xxxx pattern. String offsets are relative to the first entry.
        /// </summary>
        private Dictionary<int, string> TryParseAliasBlock(byte[] data)
        {
            var result = new Dictionary<int, string>();
            if (data == null || data.Length < 12) return result;

            // Scan for a plausible count + field type pattern
            for (int scanOff = 0; scanOff <= data.Length - 12; scanOff += 4)
            {
                int count = ByteArrayHelper.GetInt(data, scanOff);
                if (count < 1 || count > 50) continue;

                int indexSize = count * 8;
                if (scanOff + 4 + indexSize > data.Length) continue;

                // Check if the first entry looks like a field type ID (0x0B40xxxx)
                int firstTypeId = ByteArrayHelper.GetInt(data, scanOff + 4);
                if ((firstTypeId & 0xFFFF0000) != 0x0B400000) continue;

                // Looks promising — parse all entries
                // String offsets are relative to 4 bytes before the count field
                int indexOffset = scanOff + 4; // offset of first index entry
                int strBase = Math.Max(0, scanOff - 4); // offsets are relative to this
                var tempResult = new Dictionary<int, string>();
                bool valid = true;
                for (int i = 0; i < count && valid; i++)
                {
                    int entryOff = indexOffset + i * 8;
                    int typeId = ByteArrayHelper.GetInt(data, entryOff);
                    int strOff = ByteArrayHelper.GetInt(data, entryOff + 4);

                    if ((typeId & 0xFFFF0000) != 0x0B400000) { valid = false; break; }

                    int fieldIndex = typeId & 0x0000FFFF;
                    int absStrOff = strBase + strOff;
                    if (absStrOff < 0 || absStrOff >= data.Length) { valid = false; break; }

                    // Read null-terminated Unicode string
                    int strStart = absStrOff;
                    int strEnd = strStart;
                    while (strEnd + 1 < data.Length)
                    {
                        ushort ch = (ushort)(data[strEnd] | (data[strEnd + 1] << 8));
                        if (ch == 0) break;
                        strEnd += 2;
                    }
                    int strLen = strEnd - strStart;
                    if (strLen > 0)
                    {
                        string alias = System.Text.Encoding.Unicode.GetString(data, strStart, strLen);
                        if (!string.IsNullOrWhiteSpace(alias))
                            tempResult[fieldIndex] = alias;
                    }
                }

                if (valid && tempResult.Count > 0)
                    return tempResult;
            }
            return result;
        }

        /// <summary>
        /// Parse field name aliases from Props data.
        /// Format: repeated entries of (fieldTypeID: 4 bytes, aliasName: unicode null-terminated string).
        /// Returns a map of fieldIndex (lower 16 bits of typeID) → alias name.
        /// </summary>
        private Dictionary<int, string> ReadTaskFieldAliases(Props props)
        {
            var aliases = new Dictionary<int, string>();
            byte[] data = props.GetByteArray(Props.TASK_FIELD_NAME_ALIASES);
            if (data == null || data.Length < 6) return aliases;

            int offset = 0;
            while (offset + 4 < data.Length)
            {
                int fieldTypeId = ByteArrayHelper.GetInt(data, offset);
                offset += 4;

                // Read unicode string (null-terminated, 2 bytes per char)
                int strStart = offset;
                while (offset + 1 < data.Length)
                {
                    ushort ch = (ushort)(data[offset] | (data[offset + 1] << 8));
                    if (ch == 0) { offset += 2; break; }
                    offset += 2;
                }

                int strLen = offset - strStart - 2; // exclude null terminator
                if (strLen > 0)
                {
                    string alias = System.Text.Encoding.Unicode.GetString(data, strStart, strLen);
                    if (!string.IsNullOrWhiteSpace(alias))
                    {
                        int fieldIndex = fieldTypeId & 0x0000FFFF;
                        aliases[fieldIndex] = alias;
                    }
                }
            }

            return aliases;
        }

        private byte[] GetStreamDataWithDecryption(CFStorage storage, string name)
        {
            byte[] data = MppFileReader.GetStreamData(storage, name);
            if (data != null && m_inputStreamFactory != null && m_inputStreamFactory.Encrypted)
                return m_inputStreamFactory.GetData(data);
            return data;
        }
    }
}
