using System;
using System.Linq;

namespace ASTA.Classes
{
    static class StringExtensions
    {

        /// <summary>
        /// Convert date-as-string into array  int[] { 1970, 1, 1 }    .  
        /// date-as-string can be written as 'yyyyMMdd' or 'YYYY-MM-DD' or 'YYYY-MM-DD HH:MM' or 'YYYY-MM-DD HH:MM:SS' or 'YYYY-MM-DD HH:MM:SS.mmm'
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static int[] ConvertDateAsStringToIntArray(this string date)
        {
            if (date is null)
                throw new ArgumentOutOfRangeException("dateYYYYmmDD must have format 'yyyyMMdd' or 'YYYY-MM-DD' or 'YYYY-MM-DD HH:MM' or 'YYYY-MM-DD HH:MM:SS' or 'YYYY-MM-DD HH:MM:SS.mmm'");

            int[] result = new int[] { 1970, 1, 1 };
            int numberOfGropus = date.Split('-').Length;

            if (date.Contains('-') && numberOfGropus == 3)
            {
                string[] res = date.Split(' ')[0]?.Trim()?.Split('-');
                result[0] = Convert.ToInt32(res[0]);
                result[1] = Convert.ToInt32(res[1]);
                result[2] = Convert.ToInt32(res[2]);
            }
            else if (date.Length == 8)
            {
                result[0] = Convert.ToInt32(date.Remove(4));
                result[1] = Convert.ToInt32((date.Remove(0, 2)).Remove(2));
                result[2] = Convert.ToInt32(date.Remove(0, 5));
            }
            else
            {
                throw new ArgumentOutOfRangeException("dateYYYYmmDD must have format 'yyyyMMdd' or 'YYYY-MM-DD' or 'YYYY-MM-DD HH:MM' or 'YYYY-MM-DD HH:MM:SS' or 'YYYY-MM-DD HH:MM:SS.mmm'");
            }

            return result;
        }

        /// <summary>
        ///  Convert string' time to total seconds. format of time as a string can be 'HH:MM:SS' or 'HH:MM' or 'HH'
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        public static int ConvertTimeAsStringToSeconds(this string time) //time HH:MM:SS converted to decimal value
        {
            if (time is null)
            {
                return 0;
            }

            int[] result = ConvertTimeIntoStandartTimeIntArray(time);
            return (60 * 60 * result[0] + 60 * result[1] + result[2]);
        }

        /// <summary>
        /// Convert string' time 'H:M:S' to standart form 'hh:MM:ss'. 
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        public static string ConvertTimeIntoStandartTime(this string time) //time HH:MM:SS converted to decimal value
        {
            int[] result = ConvertTimeIntoStandartTimeIntArray(time);
            return string.Format("{0:d2}:{1:d2}:{2:d2}", result[0], result[1], result[2]);
        }

        /// <summary>
        /// Convert string' time 'H:M:S' to int[] {h, m, s}. format of time as a string can be 'HH:MM:SS' or 'HH:MM' or 'HH'
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        public static int[] ConvertTimeIntoStandartTimeIntArray(this string time) //time HH:MM:SS converted to decimal value
        {
            int hours = 0, minutes = 0, seconds = 0;
            int? length = time?.Split(':')?.Length;

            if (length > 0)
            {
                int.TryParse(time?.Split(':')[0], out hours);
                if (length > 1)
                {
                    int.TryParse(time?.Split(':')[1], out minutes);
                    if (length == 3)
                        int.TryParse(time?.Split(':')[2], out seconds);
                }
            }
            return new int[] { hours, minutes, seconds };
        }
    }
}
