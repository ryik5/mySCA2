using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ASTA.PersonDefinitions
{
    public class Visitors : IEnumerable
    {
        
        Visitor visitor;
      public  ObservableCollection<Visitor> visitors;

        public Visitors()
        {
            visitors = new ObservableCollection<Visitor>();
        }

        public void Add(string _fio, string _action, string _idCard, string _date, string _time, SideOfPassagePoint _sideOfPassagePoint)
        {
            visitor = new Visitor()
            {
                fio = _fio,
                idCard = _idCard,
                action = _action,
                date = _date,
                time = _time,
                sideOfPassagePoint = _sideOfPassagePoint
            };

            visitors.Add(visitor);
        }

        public ObservableCollection<Visitor> Get()
        {
            return visitors;
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)visitors).GetEnumerator();
        }

        public int Count()
        {
            return visitors.Count();
        }
    }
}
