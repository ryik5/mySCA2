namespace ASTA.Classes.People
{
    public class PeopleOutReasons
    {
        public string id { get; set; }
        public string nameReason { get; set; }
        public string visibleNameReason { get; set; }
        public int hourly { get; set; }

        public override string ToString()
        {            return $"{id}\t{nameReason}\t{visibleNameReason}\t{hourly}";        }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is PeopleOutReasons df))
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        { return ToString().GetHashCode(); }
    }

    internal class OutPerson
    {
        public string id { get; set; }//= "0"
        public string code { get; set; }//= ""
        public string date { get; set; }//= ""
        public int from { get; set; }//= 0
        public int to { get; set; }//= 0
        public int hourly { get; set; }//= 0

        public override string ToString()
        {            return $"{id}\t{code}\t{date}\t{from}\t{to}\t{hourly}";        }

        public override bool Equals(object obj)
        {
            if (obj == null||!(obj is OutPerson df))
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        {            return ToString().GetHashCode();        }
    }

    public class PeopleShift
    {
        public string code { get; set; }
        public string dayStartShift { get; set; }
        public int MoStart { get; set; }
        public int MoEnd { get; set; }
        public int TuStart { get; set; }
        public int TuEnd { get; set; }
        public int WeStart { get; set; }
        public int WeEnd { get; set; }
        public int ThStart { get; set; }
        public int ThEnd { get; set; }
        public int FrStart { get; set; }
        public int FrEnd { get; set; }
        public int SaStart { get; set; }
        public int SaEnd { get; set; }
        public int SuStart { get; set; }
        public int SuEnd { get; set; }
        public string Status { get; set; }
        public string Comment { get; set; }
      //  public new PeopleShift Get() { return this; }
    }
}