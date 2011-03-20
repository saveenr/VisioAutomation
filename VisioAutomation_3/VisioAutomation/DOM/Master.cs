using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class Master : ShapeFromMaster
    {
        public VA.Drawing.Point DropPosition { get; private set; }

        public Master(IVisio.Master master, VA.Drawing.Point dropposition) :
            base(master)
        {
            this.DropPosition = dropposition;
        }

        public Master(string master, string stencil, VA.Drawing.Point dropposition) :
            base(master, stencil)
        {
            this.DropPosition = dropposition;
        }

        public Master(IVisio.Master master, double x, double y) :
            this(master, new VA.Drawing.Point(x, y))
        {
        }
    }
}