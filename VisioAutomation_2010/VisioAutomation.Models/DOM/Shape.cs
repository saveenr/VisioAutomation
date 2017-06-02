using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Dom
{
    public class Shape : BaseShape
    {
        public MasterRef Master { get; protected set; }
        public Geometry.Point DropPosition { get; private set; }
        public Geometry.Size? DropSize { get; private set; }
        public string Name { get; set; }

        public Shape(IVisio.Master master, Geometry.Point pos)
        {
            this.Master = new MasterRef(master);
            this.DropPosition = pos;
        }

	        public Shape(IVisio.Master master, VA.Geometry.Point pos, string name)
   {
       this.Master = new MasterRef(master);
       this.DropPosition = pos;
       this.VisioShape.NameU = name;
   }

        
        public Shape(IVisio.Master master, Geometry.Rectangle rect) 
        {
            this.Master = new MasterRef(master);
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
        }

        public Shape(string mastername, string stencilname, Geometry.Point pos)
        {
            this.Master = new MasterRef(mastername, stencilname);
            this.DropPosition = pos;
        }

        public Shape(string mastername, string stencilname, Geometry.Rectangle rect) 
        {
            this.Master = new MasterRef(mastername, stencilname);
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
        }

        public Shape(IVisio.Master master, double x, double y) :
            this(master, new Geometry.Point(x, y))
        {
        }
    }
}