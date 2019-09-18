using System;
using System.Linq;

namespace ASTA
{
   public class DateTimeConvertor
    {
        private string _timeString { get; set; }
        public DateTimeConvertor() { }
        public DateTimeConvertor(string timeString)
        { _timeString = timeString; }

        public string GetStringHHMMSS(string timeString=null)
        {
            if (timeString == null && _timeString == null)
            { throw new Exception(); }

            if (timeString != null)
            {
                _timeString = timeString;
            }

          return  ConvertStringsTimeToStringHHMMSS(_timeString);
        }
        private string ConvertStringsTimeToStringHHMMSS(string time)
        {
            int h = 0;
            int m = 0;
            int s = 0;

            if (time.Contains(':'))
            {
                int.TryParse(time.Split(':')[0], out h);

                if (time.Split(':').Length > 1)
                {
                    int.TryParse(time.Split(':')[1], out m);

                    if (time.Split(':').Length > 2)
                    {
                        int.TryParse(time.Split(':')[2], out s);
                    }
                }
                return string.Format("{0:d2}:{1:d2}:{2:d2}", h, m, s);
            }
            else
            {
                return time;
            }
        }
    }
}
