using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class Master : ShapeFromMaster
    {
        public VA.Drawing.Point DropPosition { get; private set; }
        public VA.Drawing.Size? DropSize { get; private set; }

        public Master(IVisio.Master master, VA.Drawing.Point pos) :
            base(master)
        {
            this.DropPosition = pos;
        }
        
        public Master(IVisio.Master master, VA.Drawing.Rectangle rect) :
            base(master)
        {
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
        }

        public Master(string mastername, string stencilname, VA.Drawing.Point pos) :
            base(mastername, stencilname)
        {
            this.DropPosition = pos;
        }

        public Master(string mastername, string stencilname, VA.Drawing.Rectangle rect) :
            base(mastername, stencilname)
        {
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
            this.Cells.Width = rect.Size.Width;
            this.Cells.Height = rect.Size.Height;
        }

        public Master(IVisio.Master master, double x, double y) :
            this(master, new VA.Drawing.Point(x, y))
        {
        }
    }
}