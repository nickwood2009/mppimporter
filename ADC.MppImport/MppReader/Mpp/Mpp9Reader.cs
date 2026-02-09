using System;
using System.Collections.Generic;
using System.Linq;
using ADC.MppImport.MppReader.Common;
using ADC.MppImport.MppReader.Model;
using ADC.MppImport.Ole2;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Reads MPP9 files (MS Project 2003).
    /// Uses FieldMap with useTypeAsVarDataKey=false (MPP9 uses byte-based var data keys).
    /// Ported from org.mpxj.mpp.MPP9Reader
    /// </summary>
    internal class Mpp9Reader : IMppVariantReader
    {
        private const int NULL_TASK_BLOCK_SIZE = 16;

        // MPP9 default var data keys (different from MPP12/14)
        private const int TASK_NAME_VAR_KEY_DEFAULT = 1;
        private const int TASK_WBS_VAR_KEY_DEFAULT = 10;
        private const int TASK_NOTES_VAR_KEY_DEFAULT = 5;
        private const int RSC_NAME_VAR_KEY_DEFAULT = 1;
        private const int RSC_NOTES_VAR_KEY_DEFAULT = 2;

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
                throw new MppReaderException("Error reading MPP9 file: " + ex.Message, ex);
            }
        }

        private void PopulateMemberData()
        {
            byte[] propsData = MppFileReader.GetStreamData(m_root, "Props9");
            if (propsData == null)
                throw new MppReaderException("Cannot find Props9 stream");

            var props = new Props9(propsData);
            m_file.ProjectProperties.ProjectFilePath = props.GetUnicodeString(Props.PROJECT_FILE_PATH);
            m_inputStreamFactory = new DocumentInputStreamFactory(props);

            byte passwordFlag = props.GetByte(Props.PASSWORD_FLAG);
            if ((passwordFlag & 0x1) != 0 && m_reader.RespectPasswordProtection)
                throw new MppReaderException("File is password protected");

            m_projectDir = MppFileReader.GetStorage(m_root, "   19");
            if (m_projectDir == null)
            {
                m_projectDir = MppFileReader.GetStorage(m_root, "   1");
                if (m_projectDir == null)
                    throw new MppReaderException("Cannot find project directory");
            }

            byte[] projectPropsData = GetStreamDataWithDecryption(m_projectDir, "Props");
            m_projectProps = projectPropsData != null ? new Props9(projectPropsData) : new Props9(new byte[0]);

            m_file.ProjectProperties.MppFileType = 9;

            // MPP9 uses byte-based var data keys (not type-as-key)
            m_taskFieldMap = new FieldMap(false);
            m_taskFieldMap.CreateTaskFieldMap(m_projectProps);

            m_resourceFieldMap = new FieldMap(false);
            m_resourceFieldMap.CreateResourceFieldMap(m_projectProps);

            m_assignmentFieldMap = new FieldMap(false);
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
                    calVarMeta = new VarMeta9(varMetaData);
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
                    rscVarMeta = new VarMeta9(varMetaData);
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
                        resource.Name = ReadVarString(rscVarData, uniqueID, fm, (int)ResourceFieldIndex.Name, RSC_NAME_VAR_KEY_DEFAULT);
                        resource.Notes = ReadVarString(rscVarData, uniqueID, fm, (int)ResourceFieldIndex.Notes, RSC_NOTES_VAR_KEY_DEFAULT);
                    }

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
                    taskVarMeta = new VarMeta9(varMetaData);
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
                    task.ConstraintDate = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.ConstraintDate);
                    task.CreateDate = ReadFixedTimestamp(data, fm, (int)TaskFieldIndex.Created);

                    int parentUID = ReadFixedInt(data, fm, (int)TaskFieldIndex.ParentTaskUniqueID);
                    if (parentUID > 0) task.ParentTaskUniqueID = parentUID;
                    task.OutlineLevel = ReadFixedShort(data, fm, (int)TaskFieldIndex.OutlineLevel);

                    int duVal = ReadFixedShort(data, fm, (int)TaskFieldIndex.DurationUnits);
                    var du = MppUtility.GetDurationTimeUnits(duVal, properties.DefaultDurationUnits);
                    int rawDur = ReadFixedInt(data, fm, (int)TaskFieldIndex.Duration);
                    task.Duration = MppUtility.GetAdjustedDuration(properties, rawDur, du);
                    task.ActualDuration = MppUtility.GetAdjustedDuration(properties, ReadFixedInt(data, fm, (int)TaskFieldIndex.ActualDuration), du);
                    task.RemainingDuration = MppUtility.GetAdjustedDuration(properties, ReadFixedInt(data, fm, (int)TaskFieldIndex.RemainingDuration), du);

                    int ct = ReadFixedShort(data, fm, (int)TaskFieldIndex.ConstraintType);
                    if (ct >= 0 && ct <= 7) task.ConstraintType = (ConstraintType)ct;
                    task.Priority = ReadFixedShort(data, fm, (int)TaskFieldIndex.Priority);
                    task.PercentComplete = ReadFixedShort(data, fm, (int)TaskFieldIndex.PercentComplete);
                    int tt = ReadFixedShort(data, fm, (int)TaskFieldIndex.Type);
                    if (tt >= 0 && tt <= 2) task.Type = (TaskType)tt;

                    if (taskVarData != null)
                    {
                        task.Name = ReadVarString(taskVarData, uniqueID, fm, (int)TaskFieldIndex.Name, TASK_NAME_VAR_KEY_DEFAULT);
                        task.WBS = ReadVarString(taskVarData, uniqueID, fm, (int)TaskFieldIndex.WBS, TASK_WBS_VAR_KEY_DEFAULT);
                        task.Notes = ReadVarString(taskVarData, uniqueID, fm, (int)TaskFieldIndex.Notes, TASK_NOTES_VAR_KEY_DEFAULT);
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

        private string ReadVarString(Var2Data varData, int uniqueID, FieldMap fm, int fieldIndex, int defaultKey = -1)
        {
            int key = fm.GetVarDataKey(fieldIndex);
            if (key < 0) key = defaultKey >= 0 ? defaultKey : fieldIndex;
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
