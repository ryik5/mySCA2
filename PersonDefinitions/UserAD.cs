using System;
using System.Collections.Generic;

namespace ASTA.PersonDefinitions
{
    class UserAD : Employee, IComparable<UserAD>
    {
        public int id;
        public string domain;
        public string login;
        public string password;
        public string mailNickName;
        public string mail;
        public string mailServer;
        public string description;
        public string lastLogon;
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
            if (obj == null || !(obj is UserAD))
                return false;

            UserAD df = obj as UserAD;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }

        //реализация для выполнения сортировки
        int IComparable<UserAD>.CompareTo(UserAD next)
        {
            return new ADUsersComparer().Compare(this, next);
        }

        public string CompareTo(UserAD next)
        {
            return next.CompareTo(this);
        }

    }

    //additional class для выполнения сортировки
    class ADUsersComparer : IComparer<UserAD>
    {
        public int Compare(UserAD x, UserAD y)
        {
            return this.CompareTwoStaffADs(x, y);
        }

        public int CompareTwoStaffADs(UserAD x, UserAD y)
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
