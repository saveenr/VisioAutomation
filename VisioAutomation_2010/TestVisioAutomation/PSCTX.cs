using System;
using System.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SMA=System.Management.Automation;

namespace TestVisioAutomation
{
    public class PSCTX
    {
        public VisioPowerShell.Commands.New_VisioApplication visioApp = null;
        public System.Management.Automation.PowerShell powerShell;
        public System.Management.Automation.Runspaces.InitialSessionState sessionState;
        public System.Management.Automation.Runspaces.Runspace runSpace;
        public System.Management.Automation.RunspaceInvoke invoker;

        public PSCTX()
        {
            this.sessionState = SMA.Runspaces.InitialSessionState.CreateDefault();


            // Get path of where everything is executing so we can find the VisioPS.dll assembly
            var executing_assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var asm_path = System.IO.Path.GetDirectoryName(executing_assembly.GetName().CodeBase);
            var uri = new Uri(asm_path);
            var visio_ps = System.IO.Path.Combine(uri.LocalPath, "VisioPS.dll");
            var modules = new[] { visio_ps };

            // Import the latest VisioPS module into the PowerShell session
            this.sessionState.ImportPSModule(modules);
            this.runSpace = SMA.Runspaces.RunspaceFactory.CreateRunspace(this.sessionState);
            this.runSpace.Open();
            this.powerShell = SMA.PowerShell.Create();
            this.powerShell.Runspace = this.runSpace;
            this.invoker = new SMA.RunspaceInvoke(this.runSpace);
        }

        public void cleanup()
        {
            // Make sure we cleanup everything
            this.powerShell.Dispose();
            this.invoker.Dispose();
            this.runSpace.Close();
            this.invoker = null;
            this.runSpace = null;
            this.sessionState = null;
            this.powerShell = null;
        }

    }
}