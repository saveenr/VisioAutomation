using SMA = System.Management.Automation;

namespace VisioPowerShell_Tests
{
    public class VisioPS_Session : System.IDisposable
    {
        // This class should implement IDisposable because
        // it contains disposable members

        protected SMA.PowerShell PowerShell;
        protected SMA.Runspaces.InitialSessionState SessionState;
        protected SMA.Runspaces.Runspace RunSpace;
        protected SMA.RunspaceInvoke Invoker;

        public VisioPS_Session()
        {
            this.SessionState = SMA.Runspaces.InitialSessionState.CreateDefault();

            // Find the path to the assembly
            var target_assembly = typeof(VisioPowerShell.Commands.VisioCmdlet).Assembly;
            var modules = new[] { target_assembly.Location };

            // Import the latest VisioPS module into the PowerShell session
            this.SessionState.ImportPSModule(modules);
            this.RunSpace = SMA.Runspaces.RunspaceFactory.CreateRunspace(this.SessionState);
            this.RunSpace.Open();
            this.PowerShell = SMA.PowerShell.Create();
            this.PowerShell.Runspace = this.RunSpace;
            this.Invoker = new SMA.RunspaceInvoke(this.RunSpace);
        }

        public void CleanUp()
        {
            // Make sure we cleanup everything
            if (this.PowerShell != null)
            {
                this.PowerShell.Dispose();
                this.PowerShell = null;
            }
            if (this.Invoker != null)
            {
                this.Invoker.Dispose();
                this.Invoker = null;
            }
            if (this.RunSpace != null)
            {
                this.RunSpace.Close();
                this.RunSpace = null;
            }

            this.SessionState = null;
        }

        public void Dispose()
        {
            this.CleanUp();
        }
    }
}