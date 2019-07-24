using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{
    
    class InputOutputOfPerson
    {
        public int idCard { get; set; }
        public string fio { get; set; }
        public string date { get; set; }
        public string time { get; set; }
        public string action { get; set; }
        public SideOfPassagePoint sideOfPassagePoint { get; set; }

        public static InputOutputOfPersonBuilder Build()
        {
            return new InputOutputOfPersonBuilder();
        }
    }

    class InputOutputOfPersonBuilder
    {
        private InputOutputOfPerson inputOutputOfPerson;
        public InputOutputOfPersonBuilder()
        {
            inputOutputOfPerson = new InputOutputOfPersonBuilder();
        }
        public InputOutputOfPersonBuilder SetIdCard(int idCard)
        {
            inputOutputOfPerson.idCard = idCard;
            return this;
        }
        public InputOutputOfPersonBuilder SetFIO(string fio)
        {
            inputOutputOfPerson.fio = fio;
            return this;
        }
        public InputOutputOfPersonBuilder SetDate(string date)
        {
            inputOutputOfPerson.date = date;
            return this;
        }

        public InputOutputOfPersonBuilder SetTime(string time)
        {
            inputOutputOfPerson.time = time;
            return this;
        }

        public InputOutputOfPersonBuilder SetAction(string action)
        {
            inputOutputOfPerson.action = action;
            return this;
        }
        public InputOutputOfPersonBuilder SetSideOfPassagePoint(SideOfPassagePoint sideOfPassagePoint)
        {
            inputOutputOfPerson.sideOfPassagePoint = sideOfPassagePoint;
            return this;
        }

        public InputOutputOfPerson Build()
        {
            return inputOutputOfPerson;
        }

        public static implicit operator InputOutputOfPerson(InputOutputOfPersonBuilder builder)
        {
            return builder.inputOutputOfPerson;
        }
    }

    class CollectionInputsOutputs
    {
        InputOutputOfPerson inputOutputOfPerson;
        List<InputOutputOfPerson> listInputsOutputs;

        public CollectionInputsOutputs()
        {
            listInputsOutputs = new List<InputOutputOfPerson>();
        }

        public void Add(string _fio, string _action, int _idCard, string _date, string _time, SideOfPassagePoint _sideOfPassagePoint)
        {
            inputOutputOfPerson = new InputOutputOfPersonBuilder()
                .SetFIO(_fio)
                .SetIdCard(_idCard)
                .SetAction(_action)
                .SetDate(_date)
                .SetTime(_time)
                .SetSideOfPassagePoint(_sideOfPassagePoint)
                .Build();

            listInputsOutputs.Add(inputOutputOfPerson);
        }

        public List<InputOutputOfPerson> Get()
        {
            return listInputsOutputs;
        }
    }
    

    class PassagePoint
    {
        public string _idPoint;
        public string _namePoint;
        public string _connectedToServer;

        public override string ToString()
        {
            return _idPoint + "\t" + _namePoint + "\t" + _connectedToServer;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            PassagePoint df = obj as PassagePoint;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    class SideOfPassagePoint : PassagePoint
    {
        public string _direction;

        public override string ToString()
        {
            return _idPoint + "\t" + _namePoint + "\t" + _direction + "\t" + _connectedToServer;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            SideOfPassagePoint df = obj as SideOfPassagePoint;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    class SideOfPassagePointBuilder
    {
        private SideOfPassagePoint sideOfPassagePoint;
        public SideOfPassagePointBuilder()
        {
            sideOfPassagePoint = new SideOfPassagePointBuilder();
        }

        public SideOfPassagePointBuilder SetIdPoint(string idPoint)
        {
            sideOfPassagePoint._idPoint = idPoint;
            return this;
        }

        public SideOfPassagePointBuilder SetNamePoint(string namePoint)
        {
            sideOfPassagePoint._namePoint = namePoint;
            return this;
        }

        public SideOfPassagePointBuilder SetDirection(string direction)
        {
            sideOfPassagePoint._direction = direction;
            return this;
        }

        public SideOfPassagePointBuilder SetServer(string server)
        {
            sideOfPassagePoint._connectedToServer = server;
            return this;
        }

        public SideOfPassagePoint Build()
        {
            return sideOfPassagePoint;
        }

        public static implicit operator SideOfPassagePoint(SideOfPassagePointBuilder builder)
        {
            return builder.sideOfPassagePoint;
        }
    }

    class CollectionSideOfPassagePoints
    {
        SideOfPassagePoint sideOfPassagePoint;
        Dictionary<string, SideOfPassagePoint> listSideOfPassagePoints;

        public CollectionSideOfPassagePoints()
        {
            listSideOfPassagePoints = new Dictionary<string, SideOfPassagePoint>();
        }

        public void Add(string _idPoint, string _namePoint, string _direction, string _connectedToServer)
        {
            sideOfPassagePoint = new SideOfPassagePointBuilder()
                .SetIdPoint(_idPoint)
                .SetNamePoint(_namePoint)
                .SetDirection(_direction)
                .SetServer(_connectedToServer)
                .Build();

            if (_idPoint != null && !listSideOfPassagePoints.ContainsKey(_idPoint))
                listSideOfPassagePoints.Add(_idPoint, sideOfPassagePoint);
        }

        public Dictionary<string, SideOfPassagePoint> GetCollection()
        {
            return listSideOfPassagePoints;
        }

        public SideOfPassagePoint GetSideOfPassagePoint(string _idPoint)
        {
            sideOfPassagePoint = new SideOfPassagePoint();
            listSideOfPassagePoints.TryGetValue(_idPoint, out sideOfPassagePoint);

            return sideOfPassagePoint;
        }
    }



}
