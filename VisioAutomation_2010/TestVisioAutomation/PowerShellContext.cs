using System;
using SMA=System.Management.Automation;

namespace TestVisioAutomation
{
    public class PowerShellContext
    {
        public VisioPowerShell.Commands.New_VisioApplication VisioApp = null;
        public System.Management.Automation.PowerShell PowerShell;
        public System.Management.Automation.Runspaces.InitialSessionState SessionState;
        public System.Management.Automation.Runspaces.Runspace RunSpace;
        public System.Management.Automation.RunspaceInvoke Invoker;

        public PowerShellContext()
        {
            this.SessionState = SMA.Runspaces.InitialSessionState.CreateDefault();


            // Get path of where everything is executing so we can find the VisioPS.dll assembly
            var executing_assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var asm_path = System.IO.Path.GetDirectoryName(executing_assembly.GetName().CodeBase);
            var uri = new Uri(asm_path);
            var visio_ps = System.IO.Path.Combine(uri.LocalPath, "VisioPS.dll");
            var modules = new[] { visio_ps };

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
            this.PowerShell.Dispose();
            this.Invoker.Dispose();
            this.RunSpace.Close();
            this.Invoker = null;
            this.RunSpace = null;
            this.SessionState = null;
            this.PowerShell = null;
        }

    }
}