using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;
using VisioPowerShell.Commands;

namespace VisioAutomation_Tests
{
    public class VisioPSContext : PowerShellTestsSession
    {
        public VisioPSContext() : base()
        {
            
        }

        public IVisio.ShapeClass New_Visio_Container(string cont_master_name, string cont_doc)
        {
            var xdoc = this.Open_Visio_Document(cont_doc);
            var xmaster = this.Get_Visio_Master(cont_master_name,cont_doc);

            var cmd = new New_VisioContainer();
            cmd.Master = xmaster;
            var results = cmd.Invoke<IVisio.ShapeClass>();
            var shape = results.First();
            return shape ;
        }

        public List<IVisio.Shape> New_VisioShape(IVisio.MasterClass master, double[] points)
        {
            var cmd = new New_VisioShape();
            cmd.Masters = new IVisio.Master[]{ master };
            cmd.Points= points;
            var results = cmd.Invoke<List<IVisio.Shape>>();
            var shape_list = results.First();
            return shape_list;
        }

        public IVisio.MasterClass Get_Visio_Master(string rectangle, string basic_u_vss)
        {
            var doc = this.Open_Visio_Document(basic_u_vss);

            var cmd = new Get_VisioMaster();
            cmd.Name = rectangle;
            cmd.Document = doc;
            var results = cmd.Invoke<IVisio.MasterClass>();
            var master = results.First();
            return master;
        }

        public IVisio.DocumentClass Open_Visio_Document(string filename)
        {
            var cmd = new Open_VisioDocument();
            cmd.Filename = filename;
            var results = cmd.Invoke<IVisio.DocumentClass>();
            var doc = results.First();
            return doc;
        }

        public void Set_Visio_PageCells(Dictionary<string,object> dic)
        {
            var cmd = new Set_VisioShapeCell();
            cmd.Hashtable = new System.Collections.Hashtable();
            foreach (var kv in dic)
            {
                cmd.Hashtable[kv.Key] = kv.Value;
            }
            var results = cmd.Invoke();
        }

        public IVisio.DocumentClass New_Visio_Document()
        {
            var cmd = new New_VisioDocument();
            var results = cmd.Invoke<IVisio.DocumentClass>();
            var doc = results.First();
            return doc;
        }

        public IVisio.PageClass New_Visio_Page()
        {
            var cmd = new New_VisioPage();
            var results = cmd.Invoke<IVisio.PageClass>();
            var page  = results.First();
            return page;
        }


        public IVisio.ApplicationClass Get_Visio_Application()
        {
            var cmd = new Get_VisioApplication();
            var results = cmd.Invoke<IVisio.ApplicationClass>();
            var app = results.First();
            return app;
        }

        public System.Data.DataTable Get_Visio_Page_Cell( string [] cells, bool getresults, string resulttype)
        {
            var cmd = new Get_VisioPageCell();
            cmd.Cells = cells;
            cmd.GetResults = getresults;
            cmd.ResultType = VisioPowerShell.Model.ResultType.Double;
            var results = cmd.Invoke <System.Data.DataTable>();
            var dt = results.First();
            return dt;
        }

        public void Close_Visio_Application()
        {
            var cmd = new Close_VisioApplication();
            cmd.Force = true;
            var results = cmd.Invoke();
        }

    }
}