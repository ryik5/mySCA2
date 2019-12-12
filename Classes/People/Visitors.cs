using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace ASTA.Classes.People
{
    public class Visitors : INotifyPropertyChanged, IEnumerable
    {
        public ObservableRangeCollection<Visitor> collection;

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string prop)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        public Visitors()
        {
            collection = new ObservableRangeCollection<Visitor>();
        }

        public void Add(Visitor visitor, int position)
        {
            collection.Insert(position, visitor);
            OnPropertyChanged("Visitor");
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