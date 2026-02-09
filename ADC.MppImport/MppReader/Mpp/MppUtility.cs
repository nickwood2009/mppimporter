using System;
using System.Text;
using ADC.MppImport.MppReader.Common;
using ADC.MppImport.MppReader.Model;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Common utility methods for reading binary data from MPP files.
    /// Ported from org.mpxj.mpp.MPPUtility
    /// </summary>
    public static class MppUtility
    {
        private static readonly DateTime EpochDate = new DateTime(1984, 1, 1);
        private const long MsPerMinute = 60000;
        private const int DurationUnitsMask = 0xFF;

        public static void DecodeBuffer(byte[] data, byte encryptionCode)
        {
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = (byte)(data[i] ^ encryptionCode);
            }
        }

        private static readonly int[] PasswordMask =
        {
            60, 30, 48, 2, 6, 14, 8, 22, 44, 12, 38, 10, 62, 16, 34, 24
        };

        public static string DecodePassword(byte[] data, byte encryptionCode)
        {
            if (data == null || data.Length < 64)
                return null;

            DecodeBuffer(data, encryptionCode);

            var buffer = new StringBuilder();
            foreach (int index in PasswordMask)
            {
                char c = (char)data[index];
                if (c == 0) break;
                buffer.Append(c);
            }
            return buffer.ToString();
        }

        public static int GetByte(byte[] data, int offset)
        {
            return data[offset] & 0xFF;
        }

        public static long GetLong6(byte[] data, int offset)
        {
            long result = 0;
            int i = offset;
            for (int shiftBy = 0; shiftBy < 48; shiftBy += 8)
            {
                result |= ((long)(data[i] & 0xff)) << shiftBy;
                ++i;
            }
            return result;
        }

        public static double GetDouble(byte[] data, int offset)
        {
            double result = BitConverter.Int64BitsToDouble(ByteArrayHelper.GetLong(data, offset));
            if (double.IsNaN(result))
            {
                result = 0;
            }
            return result;
        }

        public static Guid? GetGUID(byte[] data, int offset)
        {
            if (data == null || data.Length <= offset + 15)
                return null;

            long long1 = 0;
            long1 |= ((long)(data[offset + 3] & 0xFF)) << 56;
            long1 |= ((long)(data[offset + 2] & 0xFF)) << 48;
            long1 |= ((long)(data[offset + 1] & 0xFF)) << 40;
            long1 |= ((long)(data[offset] & 0xFF)) << 32;
            long1 |= ((long)(data[offset + 5] & 0xFF)) << 24;
            long1 |= ((long)(data[offset + 4] & 0xFF)) << 16;
            long1 |= ((long)(data[offset + 7] & 0xFF)) << 8;
            long1 |= (long)(data[offset + 6] & 0xFF);

            long long2 = 0;
            long2 |= ((long)(data[offset + 8] & 0xFF)) << 56;
            long2 |= ((long)(data[offset + 9] & 0xFF)) << 48;
            long2 |= ((long)(data[offset + 10] & 0xFF)) << 40;
            long2 |= ((long)(data[offset + 11] & 0xFF)) << 32;
            long2 |= ((long)(data[offset + 12] & 0xFF)) << 24;
            long2 |= ((long)(data[offset + 13] & 0xFF)) << 16;
            long2 |= ((long)(data[offset + 14] & 0xFF)) << 8;
            long2 |= (long)(data[offset + 15] & 0xFF);

            if (long1 == 0 && long2 == 0)
                return null;

            // Convert from Java UUID format (two longs) to .NET Guid
            byte[] guidBytes = new byte[16];
            Array.Copy(data, offset, guidBytes, 0, 16);
            return new Guid(
                ByteArrayHelper.GetInt(data, offset),
                (short)ByteArrayHelper.GetShort(data, offset + 4),
                (short)ByteArrayHelper.GetShort(data, offset + 6),
                data[offset + 8], data[offset + 9],
                data[offset + 10], data[offset + 11],
                data[offset + 12], data[offset + 13],
                data[offset + 14], data[offset + 15]);
        }

        public static DateTime? GetDate(byte[] data, int offset)
        {
            long days = ByteArrayHelper.GetShort(data, offset);
            if (days == 65535)
                return null;
            return EpochDate.AddDays(days);
        }

        public static TimeSpan? GetTime(byte[] data, int offset)
        {
            long seconds = (ByteArrayHelper.GetShort(data, offset) / 10L) * 60L;
            if (seconds > 86399)
                seconds = seconds % 86400;
            return TimeSpan.FromSeconds(seconds);
        }

        public static long GetDurationInMs(byte[] data, int offset)
        {
            return (ByteArrayHelper.GetShort(data, offset) * MsPerMinute) / 10;
        }

        public static DateTime? GetTimestamp(byte[] data, int offset)
        {
            long days = ByteArrayHelper.GetShort(data, offset + 2);
            if (days <= 1 || days == 65535)
                return null;

            long time = ByteArrayHelper.GetShort(data, offset);
            if (time == 65535)
                time = 0;

            var result = EpochDate.AddDays(days).AddSeconds(time * 6);

            if (days < 100 && result.Second != 0)
                return null;

            return result;
        }

        public static DateTime? GetTimestampFromTenths(byte[] data, int offset)
        {
            long seconds = ((long)ByteArrayHelper.GetInt(data, offset)) * 6;
            return EpochDate.AddSeconds(seconds);
        }

        public static string GetUnicodeString(byte[] data, int offset)
        {
            int length = GetUnicodeStringLengthInBytes(data, offset);
            if (length == 0) return "";
            return Encoding.Unicode.GetString(data, offset, length);
        }

        public static string GetUnicodeString(byte[] data, int offset, int maxLength)
        {
            int length = GetUnicodeStringLengthInBytes(data, offset);
            if (maxLength > 0 && length > maxLength)
                length = maxLength;
            if (length == 0) return "";
            return Encoding.Unicode.GetString(data, offset, length);
        }

        private static int GetUnicodeStringLengthInBytes(byte[] data, int offset)
        {
            if (data == null || offset >= data.Length)
                return 0;

            int result = data.Length - offset;
            for (int loop = offset; loop < data.Length - 1; loop += 2)
            {
                if (data[loop] == 0 && data[loop + 1] == 0)
                {
                    result = loop - offset;
                    break;
                }
            }
            return result;
        }

        public static string GetString(byte[] data, int offset)
        {
            var buffer = new StringBuilder();
            for (int loop = 0; offset + loop < data.Length; loop++)
            {
                char c = (char)data[offset + loop];
                if (c == 0) break;
                buffer.Append(c);
            }
            return buffer.ToString();
        }

        public static Duration GetDuration(int value, TimeUnit type)
        {
            return GetDuration((double)value, type);
        }

        public static Duration GetDuration(double value, TimeUnit type)
        {
            double duration;
            switch (type)
            {
                case TimeUnit.Minutes:
                case TimeUnit.ElapsedMinutes:
                    duration = value / 10;
                    break;
                case TimeUnit.Hours:
                case TimeUnit.ElapsedHours:
                    duration = value / 600;
                    break;
                case TimeUnit.Days:
                    duration = value / 4800;
                    break;
                case TimeUnit.ElapsedDays:
                    duration = value / 14400;
                    break;
                case TimeUnit.Weeks:
                    duration = value / 24000;
                    break;
                case TimeUnit.ElapsedWeeks:
                    duration = value / 100800;
                    break;
                case TimeUnit.Months:
                    duration = value / 96000;
                    break;
                case TimeUnit.ElapsedMonths:
                    duration = value / 432000;
                    break;
                default:
                    duration = value;
                    break;
            }
            return Duration.GetInstance(duration, type);
        }

        public static TimeUnit GetDurationTimeUnits(int type, TimeUnit? projectDefault = null)
        {
            switch (type & DurationUnitsMask)
            {
                case 3: return TimeUnit.Minutes;
                case 4: return TimeUnit.ElapsedMinutes;
                case 5: return TimeUnit.Hours;
                case 6: return TimeUnit.ElapsedHours;
                case 7: return TimeUnit.Days;
                case 8: return TimeUnit.ElapsedDays;
                case 9: return TimeUnit.Weeks;
                case 10: return TimeUnit.ElapsedWeeks;
                case 11: return TimeUnit.Months;
                case 12: return TimeUnit.ElapsedMonths;
                case 19: return TimeUnit.Percent;
                case 20: return TimeUnit.ElapsedPercent;
                case 21: return projectDefault ?? TimeUnit.Days;
                default: return TimeUnit.Days;
            }
        }

        public static Duration GetAdjustedDuration(ProjectProperties properties, int duration, TimeUnit timeUnit)
        {
            if (duration == -1)
                return null;

            switch (timeUnit)
            {
                case TimeUnit.Days:
                {
                    double unitsPerDay = (properties.MinutesPerDay ?? 480) * 10d;
                    double totalDays = unitsPerDay != 0 ? duration / unitsPerDay : 0;
                    return Duration.GetInstance(totalDays, timeUnit);
                }
                case TimeUnit.ElapsedDays:
                {
                    double unitsPerDay = 24d * 600d;
                    return Duration.GetInstance(duration / unitsPerDay, timeUnit);
                }
                case TimeUnit.Weeks:
                {
                    double unitsPerWeek = (properties.MinutesPerWeek ?? 2400) * 10d;
                    double totalWeeks = unitsPerWeek != 0 ? duration / unitsPerWeek : 0;
                    return Duration.GetInstance(totalWeeks, timeUnit);
                }
                case TimeUnit.ElapsedWeeks:
                {
                    double unitsPerWeek = 60 * 24 * 7 * 10;
                    return Duration.GetInstance(duration / unitsPerWeek, timeUnit);
                }
                case TimeUnit.Months:
                {
                    double unitsPerMonth = (properties.MinutesPerDay ?? 480) * (properties.DaysPerMonth ?? 20) * 10d;
                    double totalMonths = unitsPerMonth != 0 ? duration / unitsPerMonth : 0;
                    return Duration.GetInstance(totalMonths, timeUnit);
                }
                case TimeUnit.ElapsedMonths:
                {
                    double unitsPerMonth = 60 * 24 * 30 * 10;
                    return Duration.GetInstance(duration / unitsPerMonth, timeUnit);
                }
                default:
                    return GetDuration(duration, timeUnit);
            }
        }

        public static TimeUnit GetWorkTimeUnits(int value)
        {
            return TimeUnitHelper.GetInstance(value - 1);
        }

        public static CurrencySymbolPosition GetSymbolPosition(int value)
        {
            switch (value)
            {
                case 1: return CurrencySymbolPosition.After;
                case 2: return CurrencySymbolPosition.BeforeWithSpace;
                case 3: return CurrencySymbolPosition.AfterWithSpace;
                default: return CurrencySymbolPosition.Before;
            }
        }

        public static string RemoveAmpersands(string name)
        {
            if (name != null && name.IndexOf('&') != -1)
            {
                var sb = new StringBuilder();
                foreach (char c in name)
                {
                    if (c != '&') sb.Append(c);
                }
                name = sb.ToString();
            }
            return name;
        }

        /// <summary>
        /// Converts a WORK-type double value (stored as milliseconds in MPP) to a Duration in hours.
        /// MPXJ WORK data type: getDouble / 60000 = hours. Values under 1000ms are treated as zero.
        /// This differs from GetDuration() which handles DURATION-type ints (tenths of a minute).
        /// </summary>
        public static Duration GetWorkDuration(double value)
        {
            if (Math.Abs(value) < 1000) return null;
            return Duration.GetInstance(value / 60000.0, TimeUnit.Hours);
        }

        public static double? GetPercentage(byte[] data, int offset)
        {
            int value = ByteArrayHelper.GetShort(data, offset);
            if (value >= 0 && value <= 100)
                return (double)value;
            return null;
        }

        public static byte[] CloneSubArray(byte[] data, int offset, int size)
        {
            byte[] newData = new byte[size];
            Array.Copy(data, offset, newData, 0, size);
            return newData;
        }
    }
}
