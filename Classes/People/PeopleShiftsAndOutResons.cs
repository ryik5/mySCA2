namespace ASTA.Classes.People
{
    public class PeopleOutReasons
    {
        public string id;
        public string nameReason;
        public string visibleNameReason;
        public int hourly;

        public override string ToString()
        {
            return $"{id}\t{nameReason}\t{visibleNameReason}\t{hourly}";
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (!(obj is PeopleOutReasons df))
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    internal class OutPerson
    {
        public string id;//= "0"
        public string code;//= ""
        public string date;//= ""
        public int from;//= 0
        public int to;//= 0
        public int hourly;//= 0

        public override string ToString()
        {
            return $"{id}\t{code}\t{date}\t{from}\t{to}\t{hourly}";
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (!(obj is OutPerson df))
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    public class PeopleShift
    {
        public string code;
        public string dayStartShift;
        public int MoStart;
        public int MoEnd;
        public int TuStart;
        public int TuEnd;
        public int WeStart;
        public int WeEnd;
        public int ThStart;
        public int ThEnd;
        public int FrStart;
        public int FrEnd;
        public int SaStart;
        public int SaEnd;
        public int SuStart;
        public int SuEnd;
        public string Status;
        public string Comment;
    }
}