namespace VisioPowerShell_Tests
{
    public class PowerShellModuleSession<T> : System.IDisposable 
    {
        // This class should implement IDisposable because
        // it contains disposable members

        protected System.Management.Automation.PowerShell PowerShell;
        protected System.Management.Automation.Runspaces.InitialSessionState SessionState;
        protected System.Management.Automation.Runspaces.Runspace RunSpace;
        protected System.Management.Automation.RunspaceInvoke Invoker;

        public PowerShellModuleSession()
        {
            this.SessionState = System.Management.Automation.Runspaces.InitialSessionState.CreateDefault();

            // Find the path to the assembly
            var target_assembly = typeof(T).Assembly;
            var modules = new[] { target_assembly.Location };

            // Import the latest VisioPS module into the PowerShell session
            this.SessionState.ImportPSModule(modules);
            this.RunSpace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(this.SessionState);
            this.RunSpace.Open();
            this.PowerShell = System.Management.Automation.PowerShell.Create();
            this.PowerShell.Runspace = this.RunSpace;
            this.Invoker = new System.Management.Automation.RunspaceInvoke(this.RunSpace);
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