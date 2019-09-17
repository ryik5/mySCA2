using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                Int32.TryParse(time.Split(':')[0], out h);

                if (time.Split(':').Length > 1)
                {
                    Int32.TryParse(time.Split(':')[1], out m);

                    if (time.Split(':').Length > 2)
                    {
                        Int32.TryParse(time.Split(':')[2], out s);
                    }
                }
                return String.Format("{0:d2}:{1:d2}:{2:d2}", h, m, s);
            }
            else
            {
                return time;
            }
        }
    }
}
