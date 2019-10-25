using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using ASTA.Classes.Common;

namespace ASTA.Classes.People
{
    public class Visitors :INotifyPropertyChanged, IEnumerable
    {

        private Visitor visitor;
        public ObservableRangeCollection<Visitor> collection;
        
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string prop)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
            }
        }

        public Visitors()
        {
            collection = new ObservableRangeCollection<Visitor>();
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
            collection.Reverse();
            collection.Add(visitor);
            collection.Reverse();
            OnPropertyChanged("Visitor");
        }

        public void Add(Visitor visitor)
        {
            collection.Add(visitor);
            OnPropertyChanged("Visitor");
        }

        public void Add(Visitor visitor, int position)
        {
            collection.Insert(position, visitor);
            OnPropertyChanged("Visitor");
        }

        public void Add(ObservableCollection<Visitor> visitors)
        {
            if (visitors?.Count > 0)
            {
                    collection.AddRange(visitors);
                    OnPropertyChanged("Visitor");
            }
        }

        public ObservableCollection<Visitor> GetCollection()
        {
            return collection;
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)collection).GetEnumerator();
        }

        public int Count()
        {
            return collection.Count();
        }
    }
}
