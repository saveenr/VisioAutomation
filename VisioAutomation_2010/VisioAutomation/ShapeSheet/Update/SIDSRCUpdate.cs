using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Update
{
    public class SIDSRCUpdate : UpdateBase
    {
        public SIDSRCUpdate() :
            base()
        {
        }

        public SIDSRCUpdate(int capacity) :
            base(capacity)
        {
        }
    }
}