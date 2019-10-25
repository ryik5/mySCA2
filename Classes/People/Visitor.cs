using System;
using System.Collections.Generic;

namespace ASTA.Classes.People
{

    public class Visitor : Person, IComparable<Visitor>//,  INotifyPropertyChanged
    {
        public string idCard { get; set; } //idCard     IdCard
        public string date { get; set; } //date of registration    DateIn
        public string time { get; set; } //time of registration    TimeIn
        public string action { get; set; }  //результат попытки прохода  ResultOfAttemptToPass

        public SideOfPassagePoint sideOfPassagePoint { get; set; }  //card reader name or id one   CardReaderName
        public Visitor() : base() { }
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

            Visitor df = obj as Visitor;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
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
            return this.CompareTwoPerson(x, y);
        }

        public int CompareTwoPerson(Visitor x, Visitor y)
        {
            string a = x.fio + x.idCard + x.date + x.time;
            string b = y.fio + y.idCard + y.date + y.time;

            string[] words = { a, b };
            Array.Sort(words);

            if (words[0] != a)
            {
                return 1;
            }
            else if (a == b)
            {
                return 0;
            }
            else
            {
                return -1;
            }
        }
    }

}
