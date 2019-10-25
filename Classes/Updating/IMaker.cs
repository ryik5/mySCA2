using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA.Classes.Updating
{
   public interface IMaker
    {
        void Make();
        void SetParameters(UpdatingParameters parameters);
        UpdatingParameters GetParameters();
    }
}
