using System;

namespace ASTA.Classes
{
    static class IntExtensions
    {
        /// <summary>
        /// Convert seconds into string 'hh:MM:ss'
        /// </summary>
        /// <param name="seconds"></param>
        /// <returns></returns>
        public static string ConvertSecondsToStringHHMMSS(this int seconds)
        {
            if (seconds == 0)
                return string.Format("{0:d2}:{1:d2}:{2:d2}", 0, 0, 0);

            int hours = seconds / 3600;
            int minutes = (seconds % 3600) / 60;
            int sec = seconds - hours * 3600 - minutes * 60;

            return string.Format("{0:d2}:{1:d2}:{2:d2}", hours, minutes, sec);
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
            result[0] = string.Format("{0:d2}", (int)ts.TotalHours);
            result[1] = string.Format("{0:d2}", (int)ts.Minutes);
            result[2] = string.Format("{0:d2}:{1:d2}", (int)ts.TotalHours, (int)ts.Minutes);

            return result;
        }
    }
}
