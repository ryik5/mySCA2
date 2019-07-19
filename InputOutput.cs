using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{
    /*
    class InputOutputOfPerson
    {
        public int idCard { get; set; }
        public string fio { get; set; }
        public string dateTime { get; set; }
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
        public InputOutputOfPersonBuilder SetDateTime(string datetime)
        {
            inputOutputOfPerson.dateTime = datetime;
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
            SetInputsOutputs();
        }

        public void InputOutputAdd(string _fio, string _action, int _idCard, string _datetime, SideOfPassagePoint _sideOfPassagePoint)
        {
            inputOutputOfPerson = new InputOutputOfPersonBuilder()
                .SetFIO(_fio)
                .SetIdCard(_idCard)
                .SetAction(_action)
                .SetDateTime(_datetime)
                .SetSideOfPassagePoint(_sideOfPassagePoint)
                .Build();

            listInputsOutputs.Add(inputOutputOfPerson);
        }

        private void SetInputsOutputs()
        {
            listInputsOutputs = new List<InputOutputOfPerson>();
        }

        public List<InputOutputOfPerson> GetInputsOutputs()
        {
            return listInputsOutputs;
        }

    }
    */

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

}
