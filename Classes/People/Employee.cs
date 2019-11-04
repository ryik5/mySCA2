using System;
using System.Collections.Generic;

namespace ASTA.Classes.People
{
    public class Employee : Person, IComparable<Employee>
    {
        public Employee() : base() { }
        public string code { get; set; }

        public override string ToString()
        {
            return fio + "\t" + code;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (!(obj is Employee df))
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        //реализация для выполнения сортировки
        int IComparable<Employee>.CompareTo(Employee next)
        {
            return new EmployeeComparer().Compare(this, next);
        }

        public string CompareTo(Employee next)
        {
            return next.CompareTo(this);
        }
    }

    public class EmployeeComparer : IComparer<Employee>
    {
        public int Compare(Employee x, Employee y)
        {
            return CompareTwoPerson(x, y);
        }

        private static int CompareTwoPerson(Employee x, Employee y)
        {
            var a = x.fio + x.code;
            var b = y.fio + y.code;

            return CompareTwoStrings.Compare(a, b);
        }
    }
}
