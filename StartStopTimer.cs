using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;

namespace ASTA
{
    class StartStopTimer
    {
        int _seconds = 0;
        string time11 = "";
        string time12 = "";
        string time21 = "";
        string time22 = "";
        string time31 = "";
        string time32 = "";
        string time41 = "";
        string time42 = "";


        public StartStopTimer(int seconds)
        {
            _seconds = seconds;
        }

        public void WaitTime1() //MiserableLoad HighPrecision
        {
            time11 = DateTime.Now.ToYYYYMMDDHHMMSSmmm();
            WaitSeconds1(_seconds);
            time12 = DateTime.Now.ToYYYYMMDDHHMMSSmmm();
        }

        public void WaitTime2()//LowLoad has problem
        {
            time21 = DateTime.Now.ToYYYYMMDDHHMMSSmmm();
            WaitSeconds2(_seconds);
            time22 = DateTime.Now.ToYYYYMMDDHHMMSSmmm();
        }

        public void WaitTime3() //MidleLoad not Precision
        {
            time31 = DateTime.Now.ToYYYYMMDDHHMMSSmmm();
            WaitSeconds3(_seconds);
            time32 = DateTime.Now.ToYYYYMMDDHHMMSSmmm();
        }

        public void WaitTime4() //HighLoad HighPrecision
        {
            time41 = DateTime.Now.ToYYYYMMDDHHMMSSmmm();
            WaitSeconds4(_seconds);
            time42 = DateTime.Now.ToYYYYMMDDHHMMSSmmm();
        }
        
        public string GetTime1()
        {
            return time11 + "\n" + time12;
        }
        public string GetTime2()
        {
            return time21 + "\n" + time22;
        }
        public string GetTime3()
        {
            return time31 + "\n" + time32;
        }
        public string GetTime4()
        {
            return time41 + "\n" + time42;
        }

        private void WaitSeconds1(int seconds) //high precision timer
        {
            System.Threading.Thread.Sleep(seconds * 1000);
        }

        private void WaitSeconds2(int seconds) //high precision timer
        {
            System.Threading.Thread.SpinWait(seconds * 26000000);
        }

        private void WaitSeconds3(int seconds) //high precision timer
        {
            //No need to risk overflowing an int
            //No need to count iterations, just set them
            //Really need to account for CPU speed, etc. though, as a
            //CPU twice as fast runs this loop twice as many times,
            //while a slow one may greatly overshoot the desired time
            //interval. Iterations too high = overshoot, too low = excessive overhead 

            System.Threading.SpinWait swt = new System.Threading.SpinWait();

            while (swt.Count < seconds * 10000)
            {
                // The NextSpinWillYield property returns true if
                // calling sw.SpinOnce() will result in yielding the
                // processor instead of simply spinning.
                swt.SpinOnce();
            }
        }

        private void WaitSeconds4(int seconds) //high precision timer
        {
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            //No need to risk overflowing an int
            while (true)
            {
                //No need to stop the stopwatch, simply "mark the lap"
                if (sw.ElapsedMilliseconds >= seconds * 1000)
                {
                    break;
                }
            }

            sw.Stop();
        }
    }
}
