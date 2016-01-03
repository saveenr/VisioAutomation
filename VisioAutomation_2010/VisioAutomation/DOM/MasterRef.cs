using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.DOM
{
    public class MasterRef
    {
        public string MasterName { get; private set; }
        public string StencilName { get; }
        public IVisio.Master VisioMaster { get; internal set; }

        public MasterRef(IVisio.Master master)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException(nameof(master));
            }

            this.VisioMaster = master;
            this.MasterName = null;
            this.StencilName = null;
        }

        public MasterRef(string mastername, string stencilname)
        {
            if (mastername == null)
            {
                throw new System.ArgumentNullException(nameof(mastername));
            }

            if (MasterRef.EndwithVSSorVSSX(mastername))
            {
                throw new AutomationException("Master name ends with .VSS or .VSSX");

                // Passing in the stencil name for the master name is a very common error.
                // so we make sure to check for it
            }

            if (this.StencilName != null)
            {
                if (!MasterRef.EndwithVSSorVSSX(stencilname))
                {                    
                    throw new AutomationException("Stencil name does not end with .VSS");

                    // Passing in the master name for the stencil name is a very common error.
                    // so we make sure to check for it

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
            return EndsWithCaseInsensitive(s,".vss") || EndsWithCaseInsensitive(s,".vssx");
        }

        private static bool EndsWithCaseInsensitive( string s, string pat)
        {
            return s.EndsWith(pat, System.StringComparison.InvariantCultureIgnoreCase);
        }
    }
}