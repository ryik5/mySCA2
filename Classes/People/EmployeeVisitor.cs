using System;
using System.Collections.Generic;

namespace ASTA.Classes.People
{
    public class EmployeeVisitor : Visitor, IComparable<EmployeeVisitor>
    {
        public string code { get; set; }
        public EmployeeVisitor() : base() { }

        public override string ToString()
        {
            return fio + "\t" + code + "\t" + idCard + "\t" + date + "\t" + time + "\t" + action + "\t" + sideOfPassagePoint._idPoint + "\t" + sideOfPassagePoint._direction;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            EmployeeVisitor df = obj as EmployeeVisitor;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        //реализация для выполнения сортировки
        int IComparable<EmployeeVisitor>.CompareTo(EmployeeVisitor next)
        {
            return new EmployeeVisitorComparer().Compare(this, next);
        }

        public string CompareTo(EmployeeVisitor next)
        {
            return next.CompareTo(this);
        }
    }

    public class EmployeeVisitorComparer : IComparer<EmployeeVisitor>
    {
        public int Compare(EmployeeVisitor x, EmployeeVisitor y)
        {
            return this.CompareTwoPerson(x, y);
        }

        public int CompareTwoPerson(EmployeeVisitor x, EmployeeVisitor y)
        {
            string a = x.fio + x.code + x.idCard + x.date + x.time;
            string b = y.fio + y.code + y.idCard + y.date + y.time;

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
