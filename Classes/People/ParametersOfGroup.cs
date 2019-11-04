using System;

namespace ASTA.Classes.People
{
   public class ParametersOfGroup
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

            ParametersOfGroup df = obj as ParametersOfGroup;
            if (df == null)
                return false;

            return ToString() == df.ToString();
        }
        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }
}
