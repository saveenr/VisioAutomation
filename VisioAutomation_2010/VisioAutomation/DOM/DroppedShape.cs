using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class DroppedShape : ShapeFromMaster
    {
        public VA.Drawing.Point DropPosition { get; private set; }
        public VA.Drawing.Size? DropSize { get; private set; }

        public DroppedShape(IVisio.Master master, VA.Drawing.Point pos) :
            base(master)
        {
            this.DropPosition = pos;
        }
        
        public DroppedShape(IVisio.Master master, VA.Drawing.Rectangle rect) :
            base(master)
        {
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
        }

        public DroppedShape(string mastername, string stencilname, VA.Drawing.Point pos) :
            base(mastername, stencilname)
        {
            this.DropPosition = pos;
        }

        public DroppedShape(string mastername, string stencilname, VA.Drawing.Rectangle rect) :
            base(mastername, stencilname)
        {
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
            this.Cells.Width = rect.Size.Width;
            this.Cells.Height = rect.Size.Height;
        }

        public DroppedShape(IVisio.Master master, double x, double y) :
            this(master, new VA.Drawing.Point(x, y))
        {
        }
    }
}