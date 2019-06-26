using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{
    //todo replace
    //using:         
    class TestTimeConvertor
    {
        public void testConvertor()
        {
            TimeConvertor timeConvertor1 = new TimeConvertor { Seconds = 115 };
            TimeStore timer = timeConvertor1;
            Console.WriteLine($"{timer.Hours}/d2:{timer.Minutes}/d2:{timer.Seconds}/d2"); // 0:1:55

            TimeConvertor timeConvertor2 = (TimeConvertor)timer;
            Console.WriteLine(timeConvertor2.Seconds);  //115
        }
    }

    class TimeStore
    {
        public int Hours { get; set; }
        public int Minutes { get; set; }
        public int Seconds { get; set; }
    }

    class TimeConvertor
    {
        public int Seconds { get; set; }

        public static implicit operator TimeConvertor(int x)
        {
            return new TimeConvertor { Seconds = x };
        }
        public static explicit operator int(TimeConvertor timeConvertor)
        {
            return timeConvertor.Seconds;
        }
        public static explicit operator TimeConvertor(TimeStore timer)
        {
            int h = timer.Hours * 3600;
            int m = timer.Minutes * 60;
            return new TimeConvertor { Seconds = h + m + timer.Seconds };
        }
        public static implicit operator TimeStore(TimeConvertor timeConvertor)
        {
            int h = timeConvertor.Seconds / 3600;
            int m = (timeConvertor.Seconds - h * 3600) / 60;
            int s = timeConvertor.Seconds - h * 3600 - m * 60;
            return new TimeStore { Hours = h, Minutes = m, Seconds = s };
        }
    }

}
