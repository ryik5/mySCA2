using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ASTA
{
    public class PassagePoint
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

    public class SideOfPassagePoint : PassagePoint
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

    public class CollectionSideOfPassagePoints : IEnumerable
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
