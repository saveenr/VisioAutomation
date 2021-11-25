using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Dom
{
    public class Shape : BaseShape
    {
        public MasterRef Master { get; protected set; }
        public VisioAutomation.Core.Point DropPosition { get; }
        public VisioAutomation.Core.Size? DropSize { get; }
        public string Name { get; set; }

        public Shape(IVisio.Master master, VisioAutomation.Core.Point pos)
        {
            this.Master = new MasterRef(master);
            this.DropPosition = pos;
        }

	        public Shape(IVisio.Master master, VA.Core.Point pos, string name)
   {
       this.Master = new MasterRef(master);
       this.DropPosition = pos;
       this.VisioShape.NameU = name;
   }

        
        public Shape(IVisio.Master master, VisioAutomation.Core.Rectangle rect) 
        {
            this.Master = new MasterRef(master);
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
        }

        public Shape(string mastername, string stencilname, VisioAutomation.Core.Point pos)
        {
            this.Master = new MasterRef(mastername, stencilname);
            this.DropPosition = pos;
        }

        public Shape(string mastername, string stencilname, VisioAutomation.Core.Rectangle rect) 
        {
            this.Master = new MasterRef(mastername, stencilname);
            this.DropPosition = rect.Center;
            this.DropSize = rect.Size;
        }

        public Shape(IVisio.Master master, double x, double y) :
            this(master, new VisioAutomation.Core.Point(x, y))
        {
        }
    }
}