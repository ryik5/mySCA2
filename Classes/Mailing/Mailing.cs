using System;

namespace ASTA.Classes
{
    internal class Mailing
    {
        public string _recipient;
        public string _groupsReport;
        public string _nameReport;
        public string _descriptionReport;
        public string _period;
        public string _status;
        public string _typeReport;
        public string _dayReport;

        public override string ToString()
        {
            return _recipient + "\t" +
                _groupsReport + "\t" + _nameReport + "\t" + _descriptionReport + "\t" + _period + "\t" +
                _status + "\t" + _typeReport + "\t" + _dayReport;
        }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is Mailing))
                return false;

            Mailing df = obj as Mailing;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    internal struct DaysOfSendingMail
    {
        public int START_OF_MONTH;
        public int MIDDLE_OF_MONTH;
        public int LAST_WORK_DAY_OF_MONTH;
        public int END_OF_MONTH;
    }

    internal class DaysWhenSendReports
    {
        DaysOfSendingMail daysOfSendingMail;

        public DaysWhenSendReports(string[] workDays, int ShiftDaysBackOfSendingFromLastWorkDay, int lastDayInMonth)
        {
            SetDaysWhenToSend(workDays, ShiftDaysBackOfSendingFromLastWorkDay, lastDayInMonth);
        }


        void SetDaysWhenToSend(string[] workDays, int ShiftDaysBackOfSendingFromLastWorkDay, int lastDayInMonth)
        {
            daysOfSendingMail = new DaysOfSendingMail
            {
                START_OF_MONTH = 1,
                END_OF_MONTH = lastDayInMonth
            };

            if (workDays.Length == 0) throw new RankException();
            foreach (string day in workDays)
            {
                if (day.Length != 10) throw new ArgumentException();
            }

            //look for last work day
            if (Int32.TryParse(workDays[workDays.Length - 1].Remove(0, 8), out int daySelected))
            {
                daysOfSendingMail.LAST_WORK_DAY_OF_MONTH = daySelected - ShiftDaysBackOfSendingFromLastWorkDay;
            }
            else
            {
                daysOfSendingMail.LAST_WORK_DAY_OF_MONTH = 28 - ShiftDaysBackOfSendingFromLastWorkDay;
            }
            daysOfSendingMail.MIDDLE_OF_MONTH = daysOfSendingMail.END_OF_MONTH / 2;
        }

        public DaysOfSendingMail GetDays()
        {
            return  daysOfSendingMail;
        }
    }
}
