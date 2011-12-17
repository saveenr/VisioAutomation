using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class ShapeFromMaster : Shape
    {
        public string MasterName { get; private set; }
        public string StencilName { get; private set; }
        public IVisio.Master MasterObject { get; internal set; }

        public ShapeFromMaster(IVisio.Master master)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException("master");
            }

            this.MasterObject = master;
            this.MasterName = null;
            this.StencilName = null;
        }

        public ShapeFromMaster(string master, string stencil)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException("master");
            }

            if (stencil == null)
            {
                throw new System.ArgumentNullException("stencil");
            }

            if (master.ToLower().EndsWith(".vss"))
            {
                throw new AutomationException("Master name ends with .VSS");
            }

            if (!stencil.ToLower().EndsWith(".vss"))
            {
                throw new AutomationException("Stencile name does not end with .VSS");
            }

            this.MasterObject = null;
            this.MasterName = master;
            this.StencilName = stencil;
        }
    }
}