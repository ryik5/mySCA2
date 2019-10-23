using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{
    class AmountMembersOfGroup
    {
        public int _amountMembers;
        public string _groupName;
        public string _emails;

        public override string ToString()
        {
            return _amountMembers + "\t" + _groupName + "\t" + _emails;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            AmountMembersOfGroup df = obj as AmountMembersOfGroup;
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
