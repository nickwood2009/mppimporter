using System;
using System.Collections.Generic;
using System.Linq;
using ADC.MppImport.MppReader.Common;
using ADC.MppImport.MppReader.Model;
using ADC.MppImport.Ole2;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Reads MPP12 files (MS Project 2007).
    /// Uses FieldMap from Props for dynamic offsets. MPP12 uses type as var data key (same as MPP14).
    /// Ported from org.mpxj.mpp.MPP12Reader
    /// </summary>
    internal class Mpp12Reader : IMppVariantReader
    {
        private const int NULL_TASK_BLOCK_SIZE = 16;

        private MppFileReader m_reader;
        private ProjectFile m_file;
        private CFStorage m_root;
        private CFStorage m_projectDir;
        private Props m_projectProps;
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
                }
            }
            catch (Exception ex)
            {
                throw new MppReaderException("Error reading MPP12 file: " + ex.Message, ex);
            }
        }

        private void PopulateMemberData()
        {
            byte[] propsData = MppFileReader.GetStreamData(m_root, "Props12");
            if (propsData == null)
                throw new MppReaderException("Cannot find Props12 stream");

            var props = new Props12(propsData);
            m_file.ProjectProperties.ProjectFilePath = props.GetUnicodeString(Props.PROJECT_FILE_PATH);
            m_inputStreamFactory = new DocumentInputStreamFactory(props);

            byte passwordFlag = props.GetByte(Props.PASSWORD_FLAG);
            if ((passwordFlag & 0x1) != 0 && m_reader.RespectPasswordProtection)
                throw new MppReaderException("File is password protected");

            m_projectDir = MppFileReader.GetStorage(m_root, "   112");
            if (m_projectDir == null)
                throw new MppReaderException("Cannot find project directory '   112'");

            byte[] projectPropsData = GetStreamDataWithDecryption(m_projectDir, "Props");
            m_projectProps = projectPropsData != null ? new Props12(projectPropsData) : new Props12(new byte[0]);

            m_file.ProjectProperties.MppFileType = 12;
            m_file.ProjectProperties.AutoFilter = props.GetBoolean(Props.AUTO_FILTER);

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
                    byte[] data = calFixedData.GetByteArrayValue(loop);
                    if (data == null || data.Length < 8) continue;

                    int calUniqueID = ByteArrayHelper.GetInt(data, 0);
                    if (calUniqueID < 1) continue;

                    var calendar = new ProjectCalendar { UniqueID = calUniqueID };
                    if (calVarData != null)
                        calendar.Name = calVarData.GetUnicodeString(calUniqueID, 1);
                    if (data.Length >= 12)
                    {
                        int parentID = ByteArrayHelper.GetInt(data, 8);
                        if (parentID > 0) calendar.ParentCalendarUniqueID = parentID;
                    }
                    m_file.Calendars.Add(calendar);
                }
            }
            catch (Exception ex) { m_file.AddIgnoredError(ex); }
        }

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

                var fm = m_resourceFieldMap;
                int itemCount = rscFixedMeta.AdjustedItemCount;
                for (int loop = 0; loop < itemCount; loop++)
                {
                    byte[] data = rscFixedData.GetByteArrayValue(loop);
                    if (data == null || data.Length < 8) continue;

                    int uniqueID = ReadFixedShort(data, fm, (int)ResourceFieldIndex.UniqueID, 0);
                    if (uniqueID < 1) continue;

                    var resource = new Resource();
                    resource.UniqueID = uniqueID;
                    resource.ID = ReadFixedInt(data, fm, (int)ResourceFieldIndex.ID, 4);

                    if (rscVarData != null)
                    {
                        resource.Name = ReadVarString(rscVarData, uniqueID, fm, (int)ResourceFieldIndex.Name);
                        resource.Initials = ReadVarString(rscVarData, uniqueID, fm, (int)ResourceFieldIndex.Initials);
                        resource.Notes = ReadVarString(rscVarData, uniqueID, fm, (int)ResourceFieldIndex.Notes);
                    }

                    resource.MaxUnits = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.MaxUnits, 8) / 100.0;
                    double stdRate = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.StandardRate, 20);
                    if (stdRate != 0) resource.StandardRate = new Rate(stdRate / 100.0, TimeUnit.Hours);
                    double otRate = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.OvertimeRate, 28);
                    if (otRate != 0) resource.OvertimeRate = new Rate(otRate / 100.0, TimeUnit.Hours);
                    double work = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.Work, 52);
                    if (work != 0) resource.Work = MppUtility.GetWorkDuration(work);
                    double cost = ReadFixedDouble(data, fm, (int)ResourceFieldIndex.Cost, 100);
                    if (cost != 0) resource.Cost = cost / 100.0;

                    resource.Active = true;
                    m_file.Resources.Add(resource);
                }
            }
            catch (Exception ex) { m_file.AddIgnoredError(ex); }
        }

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

                if (fixedMetaData == null || fixedDataBuf == null) return;

                var taskFixedMeta = new FixedMeta(fixedMetaData, 47);
                int maxSize = m_taskFieldMap.GetMaxFixedDataSize(0);
                if (maxSize < 4) maxSize = 200;
                var taskFixedData = new FixedData(taskFixedMeta, fixedDataBuf, maxSize);

                IVarMeta taskVarMeta = null;
                Var2Data taskVarData = null;
                if (varMetaData != null && var2DataBuf != null)
                {
                    taskVarMeta = new VarMeta12(varMetaData);
                    taskVarData = new Var2Data(taskVarMeta, var2DataBuf);
                }

                var properties = m_file.ProjectProperties;
                var fm = m_taskFieldMap;
                int itemCount = taskFixedMeta.AdjustedItemCount;

                var taskMap = new SortedDictionary<int, int>();
                for (int loop = itemCount - 1; loop > 2; loop--)
                {
                    byte[] data = taskFixedData.GetByteArrayValue(loop);
                    if (data == null) continue;
                    byte[] metaData = taskFixedMeta.GetByteArrayValue(loop);
                    if (metaData == null) continue;
                    int flags = ByteArrayHelper.GetInt(metaData, 0);
                    if ((flags & 0x02) != 0) continue;
                    int uid = data.Length == NULL_TASK_BLOCK_SIZE
                        ? ByteArrayHelper.GetInt(data, 0)
                        : ReadFixedInt(data, fm, (int)TaskFieldIndex.UniqueID, 0);
                    if (!taskMap.ContainsKey(uid)) taskMap[uid] = loop;
                }

                foreach (var entry in taskMap)
                {
                    int uniqueID = entry.Key;
                    byte[] data = taskFixedData.GetByteArrayValue(entry.Value);
                    if (data == null) continue;

                    if (data.Length == NULL_TASK_BLOCK_SIZE)
                    {
                        m_file.Tasks.Add(new Task { UniqueID = ByteArrayHelper.GetInt(data, 0), ID = ByteArrayHelper.GetInt(data, 4) });
                        continue;
                    }
                    if (data.Length < 8) continue;

                    var task = new Task();
                    task.UniqueID = uniqueID;
                    task.ID = ReadFixedInt(data, fm, (int)TaskFieldIndex.ID, 4);
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

                    int parentUID = ReadFixedInt(data, fm, (int)TaskFieldIndex.ParentTaskUniqueID);
                    if (parentUID > 0) task.ParentTaskUniqueID = parentUID;
                    task.OutlineLevel = ReadFixedShort(data, fm, (int)TaskFieldIndex.OutlineLevel);

                    int duVal = ReadFixedShort(data, fm, (int)TaskFieldIndex.DurationUnits);
                    var du = MppUtility.GetDurationTimeUnits(duVal, properties.DefaultDurationUnits);
                    int rawDur = ReadFixedInt(data, fm, (int)TaskFieldIndex.Duration);
                    task.Duration = MppUtility.GetAdjustedDuration(properties, rawDur, du);
                    task.ActualDuration = MppUtility.GetAdjustedDuration(properties, ReadFixedInt(data, fm, (int)TaskFieldIndex.ActualDuration), du);
                    task.RemainingDuration = MppUtility.GetAdjustedDuration(properties, ReadFixedInt(data, fm, (int)TaskFieldIndex.RemainingDuration), du);
                    task.FreeSlack = MppUtility.GetAdjustedDuration(properties, ReadFixedInt(data, fm, (int)TaskFieldIndex.FreeSlack), du);

                    int ct = ReadFixedShort(data, fm, (int)TaskFieldIndex.ConstraintType);
                    if (ct >= 0 && ct <= 7) task.ConstraintType = (ConstraintType)ct;
                    task.Priority = ReadFixedShort(data, fm, (int)TaskFieldIndex.Priority);
                    task.PercentComplete = ReadFixedShort(data, fm, (int)TaskFieldIndex.PercentComplete);
                    task.PercentWorkComplete = ReadFixedShort(data, fm, (int)TaskFieldIndex.PercentWorkComplete);
                    int tt = ReadFixedShort(data, fm, (int)TaskFieldIndex.Type);
                    if (tt >= 0 && tt <= 2) task.Type = (TaskType)tt;
                    int calID = ReadFixedInt(data, fm, (int)TaskFieldIndex.CalendarUniqueID);
                    if (calID > 0) task.CalendarUniqueID = calID;

                    double w = ReadFixedDouble(data, fm, (int)TaskFieldIndex.Work);
                    if (w != 0) task.Work = MppUtility.GetWorkDuration(w);
                    double aw = ReadFixedDouble(data, fm, (int)TaskFieldIndex.ActualWork);
                    if (aw != 0) task.ActualWork = MppUtility.GetWorkDuration(aw);
                    double rw = ReadFixedDouble(data, fm, (int)TaskFieldIndex.RemainingWork);
                    if (rw != 0) task.RemainingWork = MppUtility.GetWorkDuration(rw);
                    double co = ReadFixedDouble(data, fm, (int)TaskFieldIndex.Cost);
                    if (co != 0) task.Cost = co / 100.0;
                    double ac = ReadFixedDouble(data, fm, (int)TaskFieldIndex.ActualCost);
                    if (ac != 0) task.ActualCost = ac / 100.0;

                    if (taskVarData != null)
                    {
                        task.Name = ReadVarString(taskVarData, uniqueID, fm, (int)TaskFieldIndex.Name);
                        task.WBS = ReadVarString(taskVarData, uniqueID, fm, (int)TaskFieldIndex.WBS);
                        task.Notes = ReadVarString(taskVarData, uniqueID, fm, (int)TaskFieldIndex.Notes);
                    }

                    task.Milestone = (rawDur == 0);
                    task.Active = true;
                    if (task.Name == null && task.Start == null && task.Finish == null) continue;
                    m_file.Tasks.Add(task);
                }
            }
            catch (Exception ex) { m_file.AddIgnoredError(ex); }
        }

        private void ProcessAssignmentData()
        {
            try
            {
                CFStorage assnDir = MppFileReader.GetStorage(m_projectDir, "TBkndAssn");
                if (assnDir == null) return;
                byte[] fixedDataBuf = GetStreamDataWithDecryption(assnDir, "FixedData");
                if (fixedDataBuf == null) return;

                int maxSize = m_assignmentFieldMap.GetMaxFixedDataSize(0);
                if (maxSize < 20) maxSize = 110;
                var assnFixedData = new FixedData(maxSize, fixedDataBuf);
                var fm = m_assignmentFieldMap;

                for (int loop = 0; loop < assnFixedData.ItemCount; loop++)
                {
                    byte[] data = assnFixedData.GetByteArrayValue(loop);
                    if (data == null || data.Length < 12) continue;

                    int uid = ReadFixedInt(data, fm, (int)AssignmentFieldIndex.UniqueID, 0);
                    if (uid < 1) continue;
                    int tuid = ReadFixedInt(data, fm, (int)AssignmentFieldIndex.TaskUniqueID, 4);
                    int ruid = ReadFixedInt(data, fm, (int)AssignmentFieldIndex.ResourceUniqueID, 8);
                    if (tuid < 1) continue;

                    var a = new ResourceAssignment();
                    a.UniqueID = uid;
                    a.TaskUniqueID = tuid;
                    a.ResourceUniqueID = ruid > 0 ? ruid : (int?)null;
                    a.Start = ReadFixedTimestamp(data, fm, (int)AssignmentFieldIndex.Start, 12);
                    a.Finish = ReadFixedTimestamp(data, fm, (int)AssignmentFieldIndex.Finish, 16);
                    double u = ReadFixedDouble(data, fm, (int)AssignmentFieldIndex.Units, 46);
                    if (u != 0) a.Units = u / 100.0;
                    double w = ReadFixedDouble(data, fm, (int)AssignmentFieldIndex.Work, 54);
                    if (w != 0) a.Work = MppUtility.GetWorkDuration(w);
                    m_file.Assignments.Add(a);
                }
            }
            catch (Exception ex) { m_file.AddIgnoredError(ex); }
        }

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
            if (key < 0) key = fieldIndex;
            return varData.GetUnicodeString(uniqueID, key);
        }

        #endregion

        private byte[] GetStreamDataWithDecryption(CFStorage storage, string name)
        {
            byte[] data = MppFileReader.GetStreamData(storage, name);
            if (data != null && m_inputStreamFactory != null && m_inputStreamFactory.Encrypted)
                return m_inputStreamFactory.GetData(data);
            return data;
        }
    }
}
