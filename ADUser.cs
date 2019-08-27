using System;
using System.Collections.Generic;

namespace ASTA
{
    class ADUser : IComparable<ADUser>
    {
        public int id;
        public string domain;
        public string login;
        public string password;
        public string code;
        public string mailNickName;
        public string mail;
        public string mailServer;
        public string description;
        public string lastLogon;
        public string fio;
        public string department;
        public string stateAccount;

        //Для возможности поиска дубляжного значения
        public override string ToString()
        {
            return fio + "\t" + department + "\t" + code + "\t" +
                mail + "\t" + login + "\t" + stateAccount + "\t"
                + description + "\t" + lastLogon + "\t"
                + domain + "\t" + password;
        }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is ADUser))
                return false;

            ADUser df = obj as ADUser;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        //реализация для выполнения сортировки
        int IComparable<ADUser>.CompareTo(ADUser next)
        {
            return new ADUsersComparer().Compare(this, next);
        }

        public string CompareTo(ADUser next)
        {
            return next.CompareTo(this);
        }

    }

    //additional class для выполнения сортировки
    class ADUsersComparer : IComparer<ADUser>
    {
        public int Compare(ADUser x, ADUser y)
        {
            return this.CompareTwoStaffADs(x, y);
        }

        public int CompareTwoStaffADs(ADUser x, ADUser y)
        {
            string a = x.fio + x.login;
            string b = y.fio + y.login;

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
