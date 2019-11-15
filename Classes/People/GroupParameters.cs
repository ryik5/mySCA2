using System;

namespace ASTA.Classes.People
{
    public class GroupParameters
    {
        public int AmountMembers;
        public string GroupName;
        public string Emails;

        public override string ToString()
        {
            return $"{AmountMembers}\t{GroupName}\t{Emails}";
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            GroupParameters df = obj as GroupParameters;
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
