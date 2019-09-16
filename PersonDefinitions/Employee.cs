using System;
using System.Collections.Generic;

namespace ASTA.PersonDefinitions
{
    public class Employee : Person, IComparable<Employee>
    {
        public string code { get; set; }

        public override string ToString()
        {
            return fio + "\t" + code;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            Employee df = obj as Employee;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        //реализация для выполнения сортировки
        int IComparable<Employee>.CompareTo(Employee next)
        {
            return new PersonComparer().Compare(this, next);
        }

       public string CompareTo(Employee next)
        {
            return next.CompareTo(this);
        }
    }

  public  class PersonComparer : IComparer<Employee>
    {
        public int Compare(Employee x, Employee y)
        {
            return this.CompareTwoPerson(x, y);
        }

        public int CompareTwoPerson(Employee x, Employee y)
        {
            string a = x.fio + x.code;
            string b = y.fio + y.code;

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
