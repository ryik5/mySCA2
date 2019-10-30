using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA.Classes
{
    public class EventColorArgs
    {
        public System.Drawing.Color Color { get; }

        public EventColorArgs(System.Drawing.Color color)
        {
            Color = color;
        }
    }
}
