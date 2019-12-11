using System;
using System.Collections.Generic;

namespace ASTA.Classes.People
{
    public class ADUserFullAccount : Employee, IADUser, IComparable<ADUserFullAccount>
    {
        public int id { get; set; }
        public string Domain { get; set; }
        public string Login { get; set; }
        public string Password { get; set; }
        public string DomainControllerPath { get; set; }    // URI сервера
        public string mailNickName { get; set; }
        public string mail { get; set; }
        public string mailServer { get; set; }
        public string description { get; set; }
        public string lastLogon { get; set; }
        public string department { get; set; }
        public string stateAccount { get; set; }

        //Для возможности поиска дубляжного значения
        public override string ToString()
        {
            return $"{fio}\t{department}\t{Code}\t{mail}\t{Login}\t" +
                 $"{stateAccount}\t{description}\t{lastLogon}\t{Domain}\t{Password}";
        }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is ADUserFullAccount))
                return false;

            var df = (ADUserFullAccount)obj;
            if (df == null)
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        { return ToString().GetHashCode(); }

        //реализация для выполнения сортировки
        int IComparable<ADUserFullAccount>.CompareTo(ADUserFullAccount next)
        { return new ADUsersComparer().Compare(this, next); }

        public string CompareTo(ADUserFullAccount next)
        { return next.CompareTo(this); }
    }

    //additional class для выполнения сортировки
    public class ADUsersComparer : IComparer<ADUserFullAccount>
    {
        public int Compare(ADUserFullAccount x, ADUserFullAccount y)
        { return this.CompareTwoStaffADs(x, y); }

        public int CompareTwoStaffADs(ADUserFullAccount x, ADUserFullAccount y)
        {
            string a = x.fio + x.Login;
            string b = y.fio + y.Login;

            return CompareTwoStrings.Compare(a, b);
        }
    }
}