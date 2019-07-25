using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ASTA
{

    class InputOutputOfPerson
    {
        public string fio { get; set; }
        public string idCard { get; set; }
        public string date { get; set; }
        public string time { get; set; }
        public string action { get; set; }
        public SideOfPassagePoint sideOfPassagePoint { get; set; }
    }

    class CollectionInputsOutputs : IEnumerable
    {
        InputOutputOfPerson inputOutputOfPerson;
        List<InputOutputOfPerson> listInputsOutputs;

        public CollectionInputsOutputs()
        {
            listInputsOutputs = new List<InputOutputOfPerson>();
        }

        public void Add(string _fio, string _action, string _idCard, string _date, string _time, SideOfPassagePoint _sideOfPassagePoint)
        {
            inputOutputOfPerson = new InputOutputOfPerson()
            {
                fio = _fio,
                idCard = _idCard,
                action = _action,
                date = _date,
                time = _time,
                sideOfPassagePoint = _sideOfPassagePoint
            };

            listInputsOutputs.Add(inputOutputOfPerson);
        }

        public List<InputOutputOfPerson> Get()
        {
            return listInputsOutputs;
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)listInputsOutputs).GetEnumerator();
        }

        public int Count()
        {
            return listInputsOutputs.Count();
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
            return _namePoint + " (" + _idPoint + ")";
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

    class CollectionSideOfPassagePoints : IEnumerable
    {
        SideOfPassagePoint sideOfPassagePoint;
        Dictionary<string, SideOfPassagePoint> listSideOfPassagePoints;

        public CollectionSideOfPassagePoints()
        {
            listSideOfPassagePoints = new Dictionary<string, SideOfPassagePoint>();
        }

        public void Add(string _idPoint, string _namePoint, string _direction, string _connectedToServer)
        {
            sideOfPassagePoint = new SideOfPassagePoint()
            {
                _idPoint = _idPoint,
                _direction = _direction,
                _namePoint = _namePoint,
                _connectedToServer = _connectedToServer
            };

            if (_idPoint != null)
                if (listSideOfPassagePoints.Count == 0)
                { listSideOfPassagePoints.Add(_idPoint, sideOfPassagePoint); }
                else
                {
                    SideOfPassagePoint chkPoint;
                    if (!listSideOfPassagePoints.TryGetValue(_idPoint, out chkPoint))
                    {
                        listSideOfPassagePoints.Add(_idPoint, sideOfPassagePoint);
                    }
                }
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

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)listSideOfPassagePoints).GetEnumerator();
        }

        public int Count()
        {
            return listSideOfPassagePoints.Count();
        }
    }

}
