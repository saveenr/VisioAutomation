using SMA = System.Management.Automation;

namespace VisioPowerShell_Tests.Framework
{
    public class PowerShellSession : System.IDisposable 
    {
        protected SMA.PowerShell PowerShell;
        protected SMA.Runspaces.InitialSessionState SessionState;
        protected SMA.Runspaces.Runspace RunSpace;
        protected SMA.RunspaceInvoke Invoker;

        public PowerShellSession()
        {
            this.SessionState = SMA.Runspaces.InitialSessionState.CreateDefault();
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