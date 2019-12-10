using System;

namespace ASTA.Classes
{
    internal static class IntExtensions
    {
        /// <summary>
        /// Convert seconds into string 'hh:MM:ss'
        /// </summary>
        /// <param name="seconds"></param>
        /// <returns></returns>
        public static string ConvertSecondsToStringHHMMSS(this int seconds)
        {
            if (seconds == 0)
                return $"{0:d2}:{0:d2}:{0:d2}";

            int hours = seconds / 3600;
            int minutes = (seconds % 3600) / 60;
            int sec = seconds - hours * 3600 - minutes * 60;

            return $"{hours:d2}:{minutes:d2}:{sec:d2}";
        }

        /// <summary>
        /// Convert seconds into array hours and minutes as strings[] {"hh", "MM", "hh:MM"}
        /// </summary>
        /// <param name="seconds"></param>
        /// <returns></returns>
        public static string[] ConvertSecondsIntoStringsHHmmArray(this int seconds)
        {
            string[] result = new string[3];
            var ts = TimeSpan.FromSeconds(seconds);
            result[0] = $"{(int)ts.TotalHours:d2}";
            result[1] = $"{(int)ts.Minutes:d2}";
            result[2] = $"{(int)ts.TotalHours:d2}:{(int)ts.Minutes:d2}";

            return result;
        }
    }
}