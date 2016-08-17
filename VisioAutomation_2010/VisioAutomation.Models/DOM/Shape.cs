using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.DOM
{
    public class Shape : BaseShape
    {
        public MasterRef Master { get; protected set; }
        public Drawing.Point DropPosition { get; private set; }
        public Drawing.Size? DropSize { get; private set; }
        public string Name { get; set; }

        public Shape(IVisio.Master master, Drawing.Point pos)
        {
            this.Master = new MasterRef(master);
            this.DropPosition = pos;
        }

	        public Shape(IVisio.Master master, VA.Drawing.Point pos, string name)
   {
       this.Master = new MasterRef(master);
       this.DropPosition = pos;
       this.VisioShape.NameU = name;
   }

        
        public Shape(IVisio.Master master, Drawing.Rectangle rect) 
        {
            this.Master = new MasterRef(master);
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
        }

        public Shape(string mastername, string stencilname, Drawing.Point pos)
        {
            this.Master = new MasterRef(mastername, stencilname);
            this.DropPosition = pos;
        }

        public Shape(string mastername, string stencilname, Drawing.Rectangle rect) 
        {
            this.Master = new MasterRef(mastername, stencilname);
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
        }

        public Shape(IVisio.Master master, double x, double y) :
            this(master, new Drawing.Point(x, y))
        {
        }
    }
}