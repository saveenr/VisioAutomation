using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class DynamicConnector : ShapeFromMaster
    {
        public Shape From { get; private set; }
        public Shape To { get; private set; }

        public DynamicConnector(Shape from, Shape to, IVisio.Master master) :
            base(master)
        {
            this.From = from;
            this.To = to;
        }

        public DynamicConnector(Shape from, Shape to, string master, string stencil) :
            base(master, stencil)
        {
            this.From = from;
            this.To = to;
        }
    }
}