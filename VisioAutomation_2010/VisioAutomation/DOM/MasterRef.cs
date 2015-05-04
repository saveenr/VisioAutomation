using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.DOM
{
    using System;

    public class MasterRef
    {
        public string MasterName { get; private set; }
        public string StencilName { get; private set; }
        public IVisio.Master VisioMaster { get; internal set; }

        public MasterRef(IVisio.Master master)
        {
            if (master == null)
            {
                throw new ArgumentNullException("master");
            }

            this.VisioMaster = master;
            this.MasterName = null;
            this.StencilName = null;
        }

        public MasterRef(string mastername, string stencilname)
        {
            if (mastername == null)
            {
                throw new ArgumentNullException("mastername");
            }


            if (EndwithVSSorVSSX(mastername))
            {
                throw new AutomationException("Master name ends with .VSS or .VSSX");
            }

            if (this.StencilName != null)
            {
                if (!EndwithVSSorVSSX(stencilname))
                {                    
                    throw new AutomationException("Stencil name does not end with .VSS");
                }
            }
            else
            {
                // Stencil names are allowed to be null. In this case 
                // it means look for the stencil in the active document
            }

            this.VisioMaster = null;
            this.MasterName = mastername;
            this.StencilName = stencilname;
        }

        private static bool EndwithVSSorVSSX(string s)
        {
            return s.EndsWith(".vss", StringComparison.InvariantCultureIgnoreCase) || s.EndsWith(".vssx", StringComparison.InvariantCultureIgnoreCase);
        }
    }
}