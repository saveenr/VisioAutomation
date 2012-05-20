using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.DOM
{
    public class Shape : BaseShape
    {
        public VA.DOM.MasterRef Master { get; protected set; }
        public VA.Drawing.Point DropPosition { get; private set; }
        public VA.Drawing.Size? DropSize { get; private set; }

        public Shape(IVisio.Master master, VA.Drawing.Point pos)
        {
            this.Master = new VA.DOM.MasterRef(master);
            this.DropPosition = pos;
        }
        
        public Shape(IVisio.Master master, VA.Drawing.Rectangle rect) 
        {
            this.Master = new VA.DOM.MasterRef(master);
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
        }

        public Shape(string mastername, string stencilname, VA.Drawing.Point pos)
        {
            this.Master = new VA.DOM.MasterRef(mastername, stencilname);
            this.DropPosition = pos;
        }

        public Shape(string mastername, string stencilname, VA.Drawing.Rectangle rect) 
        {
            this.Master = new VA.DOM.MasterRef(mastername, stencilname);
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
            this.Cells.Width = rect.Size.Width;
            this.Cells.Height = rect.Size.Height;
        }

        public Shape(IVisio.Master master, double x, double y) :
            this(master, new VA.Drawing.Point(x, y))
        {
        }
    }
}