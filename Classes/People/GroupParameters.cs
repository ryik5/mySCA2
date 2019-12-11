namespace ASTA.Classes.People
{
    public class GroupParameters
    {
        public int AmountMembers { get; set; }
        public string GroupName { get; set; }
        public string Emails { get; set; }

        public override string ToString()
        { return $"{AmountMembers}\t{GroupName}\t{Emails}"; }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is GroupParameters df))
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        { return ToString().GetHashCode(); }
    }
}