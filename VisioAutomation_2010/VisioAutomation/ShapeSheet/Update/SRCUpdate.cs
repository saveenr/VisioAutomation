using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Update
{
    public class SRCUpdate : UpdateBase
    {
        public SRCUpdate() :
            base()
        {
        }

        public SRCUpdate(int capacity) :
            base(capacity)
        {
        }
    }
}