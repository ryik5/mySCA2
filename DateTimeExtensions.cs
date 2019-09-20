using System;
using System.Linq;

namespace ASTA
{

    static class DateTimeExtensions
    {
        public static DateTime LastDayOfMonth(this DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
        }

        public static DateTime FirstDayOfMonth(this DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
        }

        public static string ToMonthName(this DateTime dateTime)
        {
            return System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(dateTime.Month);
        }

        public static string ToShortMonthName(this DateTime dateTime)
        {
            return System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(dateTime.Month);
        }

        public static string ToYYYYMMDD(this DateTime dateTime)
        {
            return dateTime.ToString("yyyy-MM-dd");
        }

        public static int[] ToIntYYYYMMDD(this DateTime dateTime)
        {
            return new int[] { dateTime.Year, dateTime.Month, dateTime.Day };
        }

        public static string ToYYYYMMDDHHMM(this DateTime dateTime)
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm");
        }

        public static string ToYYYYMMDDHHMMSS(this DateTime dateTime)
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }
        public static string ToYYYYMMDDHHMMSSmmm(this DateTime dateTime)
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        }

    }

}
