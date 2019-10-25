using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA.Classes
{
    public class AccountEventBoolArgs
    {
        public bool Status { get; }
        public System.Drawing.Color Color { get; }

        public AccountEventBoolArgs(bool status, System.Drawing.Color color)
        {
            Color = color;
            Status = status;
        }
    }
}
