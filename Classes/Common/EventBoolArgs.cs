using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA.Classes
{
  public  class EventBoolArgs
    {
        public bool Status { get; }
        public EventBoolArgs(bool status)
        {
            Status = status;
        }
    }
}
