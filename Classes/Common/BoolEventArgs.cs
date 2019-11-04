
using System;

namespace ASTA.Classes
{
  public  class BoolEventArgs : EventArgs
    {
        public bool Status { get; private set; }
        public BoolEventArgs(bool status)
        {
            Status = status;
        }
    }
}
