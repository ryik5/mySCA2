using System;
using System.Collections.Generic;

namespace ASTA.Classes.People
{
    public class Employee : IComparable<Employee>
    {
        public string fio { get; set; }

        public string Code { get; set; }

        public Employee() { }

       // public new Employee Get() { return this; }

        public override string ToString()
        { return $"{fio}\t{Code}"; }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is Employee df))
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        { return ToString().GetHashCode(); }

        //реализация для выполнения сортировки
        int IComparable<Employee>.CompareTo(Employee next)
        { return new EmployeeComparer().Compare(this, next); }

        public string CompareTo(Employee next)
        { return next.CompareTo(this); }
    }

    public class EmployeeComparer : IComparer<Employee>
    {
        public int Compare(Employee x, Employee y)
        { return CompareTwoPerson(x, y); }

        private static int CompareTwoPerson(Employee x, Employee y)
        {
            var a = x.fio + x.Code;
            var b = y.fio + y.Code;

            return CompareTwoStrings.Compare(a, b);
        }
    }
}