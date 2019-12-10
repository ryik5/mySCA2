using System;
using System.Collections.Generic;

namespace ASTA.Classes.People
{
    public class Visitor : IComparable<Visitor>//,  INotifyPropertyChanged
    {
        public string fio { get; set; }

        public string idCard { get; set; } //idCard     IdCard
        public string date { get; set; } //date of registration    DateIn
        public string time { get; set; } //time of registration    TimeIn
        public string action { get; set; }  //результат попытки прохода  ResultOfAttemptToPass

        public SideOfPassagePoint sideOfPassagePoint { get; set; }  //card reader name or id one   CardReaderName

        public Visitor()
        {
        }

        public Visitor(string _fio, string _idCard, string _date, string _time, string _action, SideOfPassagePoint _sideOfPassagePoint)
        {
            fio = _fio;
            idCard = _idCard;
            date = _date;
            time = _time;
            action = _action;
            sideOfPassagePoint = _sideOfPassagePoint;
        }

        public override string ToString()
        {
            return fio + "\t" + idCard + "\t" + date + "\t" + time + "\t" + action + "\t" + sideOfPassagePoint._idPoint + "\t" + sideOfPassagePoint._direction;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (!(obj is Visitor df))
                return false;

            return ToString().Equals(df.ToString());
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        //реализация для выполнения сортировки
        int IComparable<Visitor>.CompareTo(Visitor next)
        {
            return new VisitorComparer().Compare(this, next);
        }

        public string CompareTo(Visitor next)
        {
            return next.CompareTo(this);
        }
    }

    public class VisitorComparer : IComparer<Visitor>
    {
        public int Compare(Visitor x, Visitor y)
        {
            return CompareTwoPerson(x, y);
        }

        public int CompareTwoPerson(Visitor x, Visitor y)
        {
            string a = x.fio + x.idCard + x.date + x.time;
            string b = y.fio + y.idCard + y.date + y.time;

            return CompareTwoStrings.Compare(a, b);
        }
    }
}