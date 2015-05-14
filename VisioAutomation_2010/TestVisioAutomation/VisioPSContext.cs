using SMA=System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;

namespace TestVisioAutomation
{
    public class VisioPSContext : PowerShellSession
    {
        public VisioPSContext() : base()
        {
            
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
            var cmd = new VisioPowerShell.Commands.New_VisioDocument();
            var results = cmd.Invoke<IVisio.DocumentClass>();
            var doc = results.First();
            return doc;
        }

        public IVisio.PageClass New_Visio_Page()
        {
            var cmd = new VisioPowerShell.Commands.New_VisioPage();
            var results = cmd.Invoke<IVisio.PageClass>();
            var page  = results.First();
            return page;
        }


        public IVisio.ApplicationClass Get_Visio_Application()
        {
            var cmd = new VisioPowerShell.Commands.Get_VisioApplication();
            var results = cmd.Invoke<IVisio.ApplicationClass>();
            var app = results.First();
            return app;
        }

        public System.Data.DataTable Get_Visio_Page_Cell( string [] Cells, bool GetResults, string ResultType)
        {
            var cmd = new VisioPowerShell.Commands.Get_VisioPageCell();
            cmd.Cells = Cells;
            cmd.GetResults = GetResults;
            cmd.ResultType = VisioPowerShell.ResultType.Double;
            var results = cmd.Invoke <System.Data.DataTable>();
            var dt = results.First();
            return dt;
        }

        public void Close_Visio_Application()
        {
            var cmd = new VisioPowerShell.Commands.Close_VisioApplication();
            cmd.Force = true;
            var results = cmd.Invoke();
        }

    }
}