using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class DynamicConnector : DroppedShape
    {
        public Shape From { get; private set; }
        public Shape To { get; private set; }
        
        public DynamicConnector(Shape from, Shape to, IVisio.Master master) :
            base(master,-3,-3)
        {
            this.Master = new VA.DOM.MasterRef(master);
            this.From = from;
            this.To = to;
        }

        public DynamicConnector(Shape from, Shape to, string mastername, string stencilname) :
            base(mastername,stencilname, new VA.Drawing.Point(-3,-3) )
        {
            this.Master = new VA.DOM.MasterRef(mastername, stencilname);
            this.From = from;
            this.To = to;
        }
    }
}