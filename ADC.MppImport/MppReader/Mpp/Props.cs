using System;
using System.Collections.Generic;
using System.Text;
using ADC.MppImport.MppReader.Common;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Base class for Props file readers. Props files contain key-value property data.
    /// Ported from org.mpxj.mpp.Props
    /// </summary>
    internal class Props
    {
        protected readonly SortedDictionary<int, byte[]> m_map = new SortedDictionary<int, byte[]>();

        // Property key constants
        public static readonly int PROJECT_START_DATE = 37748738;
        public static readonly int PROJECT_FINISH_DATE = 37748739;
        public static readonly int SCHEDULE_FROM = 37748740;
        public static readonly int RESOURCE_POOL = 37748747;
        public static readonly int DEFAULT_CALENDAR_NAME = 37748750;
        public static readonly int CURRENCY_SYMBOL = 37748752;
        public static readonly int CURRENCY_PLACEMENT = 37748753;
        public static readonly int CURRENCY_DIGITS = 37748754;
        public static readonly int CRITICAL_SLACK_LIMIT = 37748756;
        public static readonly int DURATION_UNITS = 37748757;
        public static readonly int WORK_UNITS = 37748758;
        public static readonly int TASK_UPDATES_RESOURCE = 37748761;
        public static readonly int SPLIT_TASKS = 37748762;
        public static readonly int START_TIME = 37748764;
        public static readonly int MINUTES_PER_DAY = 37748765;
        public static readonly int MINUTES_PER_WEEK = 37748766;
        public static readonly int STANDARD_RATE = 37748767;
        public static readonly int OVERTIME_RATE = 37748768;
        public static readonly int END_TIME = 37748769;
        public static readonly int WEEK_START_DAY = 37748773;
        public static readonly int GUID = 37748777;
        public static readonly int FISCAL_YEAR_START_MONTH = 37748780;
        public static readonly int DEFAULT_TASK_TYPE = 37748785;
        public static readonly int HONOR_CONSTRAINTS = 37748794;
        public static readonly int FISCAL_YEAR_START = 37748801;
        public static readonly int EDITABLE_ACTUAL_COSTS = 37748802;
        public static readonly int DAYS_PER_MONTH = 37753743;
        public static readonly int CURRENCY_CODE = 37753787;
        public static readonly int NEW_TASKS_ARE_MANUAL = 37753800;
        public static readonly int MULTIPLE_CRITICAL_PATHS = 37748793;
        public static readonly int TASK_FIELD_NAME_ALIASES = 1048577;
        public static readonly int RESOURCE_FIELD_NAME_ALIASES = 1048578;
        public static readonly int PASSWORD_FLAG = 893386752;
        public static readonly int PROTECTION_PASSWORD_HASH = 893386756;
        public static readonly int WRITE_RESERVATION_PASSWORD_HASH = 893386757;
        public static readonly int ENCRYPTION_CODE = 893386759;
        public static readonly int STATUS_DATE = 37748805;
        public static readonly int SUBPROJECT_COUNT = 37748868;
        public static readonly int SUBPROJECT_DATA = 37748898;
        public static readonly int SUBPROJECT_TASK_COUNT = 37748900;
        public static readonly int DEFAULT_CALENDAR_HOURS = 37753736;
        public static readonly int TASK_FIELD_ATTRIBUTES = 37753744;
        public static readonly int FONT_BASES = 54525952;
        public static readonly int AUTO_FILTER = 893386767;
        public static readonly int PROJECT_FILE_PATH = 893386760;
        public static readonly int HYPERLINK_BASE = 37748810;
        public static readonly int RESOURCE_CREATION_DATE = 205521219;
        public static readonly int SHOW_PROJECT_SUMMARY_TASK = 54525961;
        public static readonly int TASK_FIELD_MAP = 131092;
        public static readonly int TASK_FIELD_MAP2 = 50331668;
        public static readonly int ENTERPRISE_CUSTOM_FIELD_MAP = 37753797;
        public static readonly int RESOURCE_FIELD_MAP = 131093;
        public static readonly int RESOURCE_FIELD_MAP2 = 50331669;
        public static readonly int RELATION_FIELD_MAP = 131094;
        public static readonly int ASSIGNMENT_FIELD_MAP = 131095;
        public static readonly int ASSIGNMENT_FIELD_MAP2 = 50331671;
        public static readonly int BASELINE_CALENDAR_NAME = 37753747;
        public static readonly int BASELINE_DATE = 37753749;
        public static readonly int BASELINE1_DATE = 37753750;
        public static readonly int BASELINE2_DATE = 37753751;
        public static readonly int BASELINE3_DATE = 37753752;
        public static readonly int BASELINE4_DATE = 37753753;
        public static readonly int BASELINE5_DATE = 37753754;
        public static readonly int BASELINE6_DATE = 37753755;
        public static readonly int BASELINE7_DATE = 37753756;
        public static readonly int BASELINE8_DATE = 37753757;
        public static readonly int BASELINE9_DATE = 37753758;
        public static readonly int BASELINE10_DATE = 37753759;
        public static readonly int CUSTOM_FIELDS = 71303169;

        public byte[] GetByteArray(int type)
        {
            m_map.TryGetValue(type, out byte[] result);
            return result;
        }

        public byte GetByte(int type)
        {
            if (m_map.TryGetValue(type, out byte[] item) && item != null && item.Length > 0)
                return item[0];
            return 0;
        }

        public int GetShort(int type)
        {
            if (m_map.TryGetValue(type, out byte[] item) && item != null && item.Length >= 2)
                return ByteArrayHelper.GetShort(item, 0);
            return 0;
        }

        public int GetInt(int type)
        {
            if (m_map.TryGetValue(type, out byte[] item) && item != null && item.Length >= 4)
                return ByteArrayHelper.GetInt(item, 0);
            return 0;
        }

        public double GetDouble(int type)
        {
            if (m_map.TryGetValue(type, out byte[] item) && item != null && item.Length >= 8)
                return MppUtility.GetDouble(item, 0);
            return 0;
        }

        public TimeSpan? GetTime(int type)
        {
            if (m_map.TryGetValue(type, out byte[] item) && item != null && item.Length >= 2)
                return MppUtility.GetTime(item, 0);
            return null;
        }

        public DateTime? GetTimestamp(int type)
        {
            if (m_map.TryGetValue(type, out byte[] item) && item != null && item.Length >= 4)
                return MppUtility.GetTimestamp(item, 0);
            return null;
        }

        public bool GetBoolean(int type)
        {
            if (m_map.TryGetValue(type, out byte[] item) && item != null && item.Length >= 2)
                return ByteArrayHelper.GetShort(item, 0) != 0;
            return false;
        }

        public string GetUnicodeString(int type)
        {
            if (m_map.TryGetValue(type, out byte[] item) && item != null)
                return MppUtility.GetUnicodeString(item, 0);
            return null;
        }

        public DateTime? GetDate(int type)
        {
            if (m_map.TryGetValue(type, out byte[] item) && item != null && item.Length >= 2)
                return MppUtility.GetDate(item, 0);
            return null;
        }

        public Guid? GetUUID(int type)
        {
            if (!m_map.TryGetValue(type, out byte[] item) || item == null)
                return null;

            if (item.Length > 16)
            {
                string value = MppUtility.GetUnicodeString(item, 0, 76);
                if (value.Length == 38 && value[0] == '{' && value[37] == '}')
                {
                    return Guid.Parse(value.Substring(1, 36));
                }
            }

            return MppUtility.GetGUID(item, 0);
        }

        public ICollection<int> KeySet()
        {
            return m_map.Keys;
        }
    }
}
