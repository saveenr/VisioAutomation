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

        public ShapeFromMaster(string mastername, string stencilname)
        {
            if (mastername == null)
            {
                throw new System.ArgumentNullException("mastername");
            }

            if (stencilname == null)
            {
                throw new System.ArgumentNullException("stencilname");
            }

            if (mastername.ToLower().EndsWith(".vss"))
            {
                throw new AutomationException("Master name ends with .VSS");
            }

            if (!stencilname.ToLower().EndsWith(".vss"))
            {
                throw new AutomationException("Stencile name does not end with .VSS");
            }

            this.MasterObject = null;
            this.MasterName = mastername;
            this.StencilName = stencilname;
        }
    }
}