using System;

namespace ASTA.Classes.People
{    
    class OutReasons
    {
        public string _id;
        public string _name;
        public string _visibleName;
        public int _hourly;

        public override string ToString()
        {
            return _id + "\t" + _name + "\t" + _visibleName + "\t" + _hourly;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            OutReasons df = obj as OutReasons;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    class OutPerson
    {
        public string _reason_id;//= "0"
        public string _nav;//= ""
        public string _date;//= ""
        public int _from;//= 0
        public int _to;//= 0
        public int _hourly;//= 0

        public override string ToString()
        {
            return _reason_id + "\t" + _nav + "\t" + _date + "\t" + _from + "\t" + _to + "\t" + _hourly;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            OutPerson df = obj as OutPerson;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    struct PeopleShift
    {
        public string _nav;
        public string _dayStartShift;
        public int _MoStart;
        public int _MoEnd;
        public int _TuStart;
        public int _TuEnd;
        public int _WeStart;
        public int _WeEnd;
        public int _ThStart;
        public int _ThEnd;
        public int _FrStart;
        public int _FrEnd;
        public int _SaStart;
        public int _SaEnd;
        public int _SuStart;
        public int _SuEnd;
        public string _Status;
        public string _Comment;
    }

}
