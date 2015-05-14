using System;
using SMA=System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace TestVisioAutomation
{

    public class PowerShellContext
    {
        private System.Management.Automation.PowerShell PowerShell;
        private System.Management.Automation.Runspaces.InitialSessionState SessionState;
        private System.Management.Automation.Runspaces.Runspace RunSpace;
        private System.Management.Automation.RunspaceInvoke Invoker;

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

        public IVisio.ShapeClass New_Visio_Container(string cont_master_name, string cont_doc)
        {
            var cmd = string.Format("New-VisioContainer -Master (Get-VisioMaster \"{0}\" (Open-VisioDocument \"{1}\"))", cont_master_name, cont_doc);
            var results = this.Invoker.Invoke(cmd);
            var shape = (IVisio.ShapeClass)results[0].BaseObject;
            return shape;
        }

        public List<IVisio.Shape> New_VisioShape(IVisio.MasterClass master, double[] points)
        {
            var pipeline = this.RunSpace.CreatePipeline();
            var cmd = new SMA.Runspaces.Command(@"New-VisioShape");
            cmd.AddParameter("Master", master);
            cmd.AddParameter("Points", points);
            pipeline.Commands.Add(cmd);
            var results = pipeline.Invoke();
            var shapes = (List<IVisio.Shape>)results[0].BaseObject;
            return shapes;
        }

        public IVisio.MasterClass Get_Visio_Master(string rectangle, string basic_u_vss)
        {
            var cmd = string.Format("(Get-VisioMaster \"{0}\" (Open-VisioDocument \"{1}\"))", rectangle, basic_u_vss);
            var results = this.Invoker.Invoke(cmd);
            var master = (IVisio.MasterClass)results[0].BaseObject;
            return master;
        }

        public IVisio.DocumentClass New_Visio_Document()
        {
            var results = this.Invoker.Invoke("New-VisioDocument");
            var doc = (IVisio.DocumentClass)results[0].BaseObject;
            return doc;
        }

        public IVisio.PageClass New_Visio_Page()
        {
            var results = this.Invoker.Invoke("New-VisioPage");
            var page = (IVisio.PageClass)results[0].BaseObject;
            return page;
        }


        public IVisio.ApplicationClass Get_Visio_Application()
        {
            var app_0 = this.Invoker.Invoke("Get-VisioApplication");
            var app = (IVisio.ApplicationClass)app_0[0].BaseObject;
            return app;
        }

        public System.Data.DataTable Get_Visio_Page_Cell( string [] Cells, bool GetResults, string ResultType)
        {
            var pipeline = this.RunSpace.CreatePipeline();
            var cmd = new SMA.Runspaces.Command(@"Get-VisioPageCell");
            cmd.AddParameter("Cells", Cells);
            cmd.AddParameter("GetResults", GetResults);
            cmd.AddParameter("ResultType", ResultType);
            pipeline.Commands.Add(cmd);
            var results = pipeline.Invoke();
            var dt = (System.Data.DataTable)results[0].BaseObject;
            return dt;
        }

        public void Close_Visio_Application()
        {
            this.Invoker.Invoke("Close-VisioApplication -Force");
        }

    }
}