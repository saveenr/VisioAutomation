using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.DOM
{
    public class Connector : Shape
    {
        public BaseShape From { get; private set; }
        public BaseShape To { get; private set; }
        
        public Connector(BaseShape from, BaseShape to, IVisio.Master master) :
            base(master,-3,-3)
        {
            this.Master = new VA.DOM.MasterRef(master);
            this.From = from;
            this.To = to;
        }

        public Connector(BaseShape from, BaseShape to, string mastername, string stencilname) :
            base(mastername,stencilname, new VA.Drawing.Point(-3,-3) )
        {
            this.Master = new VA.DOM.MasterRef(mastername, stencilname);
            this.From = from;
            this.To = to;
        }
    }
}