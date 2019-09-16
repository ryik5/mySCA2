
using System.ComponentModel;

namespace ASTA.PersonDefinitions
{

    public class Visitor : Person//,  INotifyPropertyChanged
    {
        public string idCard { get; set; } //idCard     IdCard
        public string date { get; set; } //date of registration    DateIn
        public string time { get; set; } //time of registration    TimeIn
        public string action { get; set; }  //результат попытки прохода  ResultOfAttemptToPass

        public SideOfPassagePoint sideOfPassagePoint { get; set; }  //card reader name or id one   CardReaderName

       // public event PropertyChangedEventHandler PropertyChanged;
    }
}
