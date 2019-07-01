using System;
using System.Collections.Generic;

namespace ASTA
{

    interface IPerson
    {
        string FIO { get; set; }
        string NAV { get; set; }
    }

    class PersonClasses : IPerson, IComparable<PersonClasses>
    {
        public string FIO { get; set; }
        public string NAV { get; set; }

        public override string ToString()
        {
            return FIO + "\t" + NAV;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            PersonClasses df = obj as PersonClasses;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        //реализация для выполнения сортировки
        int IComparable<PersonClasses>.CompareTo(PersonClasses next)
        {
            return new PersonComparer().Compare(this, next);
        }

        string CompareTo(PersonClasses next)
        {
            return next.CompareTo(this);
        }
    }

    class PersonComparer : IComparer<PersonClasses>
    {
        public int Compare(PersonClasses x, PersonClasses y)
        {
            return this.CompareTwoPerson(x, y);
        }

        public int CompareTwoPerson(PersonClasses x, PersonClasses y)
        {
            string a = x.FIO + x.NAV;
            string b = y.FIO + y.NAV;

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

    class PersonFull : IPerson
    {
        public int idCard;//= 0
        public string FIO { get; set; }//= ""
        public string NAV { get; set; }//= ""
        public string Department;//= ""
        public string DepartmentId;//= ""
        public string DepartmentBossCode;//= ""
        public string PositionInDepartment;//= ""
        public string GroupPerson;//= ""
        public string City;//= ""
        public int ControlInSeconds;//= 32460
        public int ControlOutSeconds;// =64800
        public string ControlInHHMM;//= "09:00"
        public string ControlOutHHMM;//= "18:00"
        public string Shift;//= ""
        public string Comment;//= ""

        public override string ToString()
        {
            return idCard + "\t" + FIO + "\t" + NAV + "\t" + this.Department + "\t" + DepartmentId + "\t" + DepartmentBossCode + "\t" +
                PositionInDepartment + "\t" + GroupPerson + "\t" + City + "\t" +
                ControlInSeconds + "\t" + ControlOutSeconds + "\t" + ControlInHHMM + "\t" + ControlOutHHMM + "\t" +
                Shift + "\t" + Comment;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            PersonFull df = obj as PersonFull;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

}
