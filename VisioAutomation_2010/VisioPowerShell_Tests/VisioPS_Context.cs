using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;

namespace VisioPowerShell_Tests
{
    public class VisioPS_Context : VisioPS_TestSession
    {
        public VisioPS_Context() : base()
        {
            
        }

        public IVisio.ShapeClass New_VisioContainer(string cont_master_name, string cont_doc)
        {
            var xdoc = this.Open_VisioDocument(cont_doc);
            var xmaster = this.Get_VisioMaster(cont_master_name,cont_doc);

            var cmd = new VisioPowerShell.Commands.NewVisioContainer();
            cmd.Master = xmaster;
            var results = cmd.Invoke<IVisio.ShapeClass>();
            var shape = results.First();
            return shape ;
        }

        public List<IVisio.Shape> New_VisioShape(IVisio.MasterClass master, double[] points)
        {
            var cmd = new VisioPowerShell.Commands.NewVisioShape();
            cmd.Masters = new IVisio.Master[]{ master };
            cmd.Points= points;
            var results = cmd.Invoke<List<IVisio.Shape>>();
            var shape_list = results.First();
            return shape_list;
        }

        public IVisio.MasterClass Get_VisioMaster(string rectangle, string basic_u_vss)
        {
            var doc = this.Open_VisioDocument(basic_u_vss);

            var cmd = new VisioPowerShell.Commands.GetVisioMaster();
            cmd.Name = rectangle;
            cmd.Document = doc;
            var results = cmd.Invoke<IVisio.MasterClass>();
            var master = results.First();
            return master;
        }

        public IVisio.DocumentClass Open_VisioDocument(string filename)
        {
            var cmd = new VisioPowerShell.Commands.OpenVisioDocument();
            cmd.Filename = filename;
            var results = cmd.Invoke<IVisio.DocumentClass>();
            var doc = results.First();
            return doc;
        }

        public IVisio.DocumentClass New_VisioDocument()
        {
            var cmd = new VisioPowerShell.Commands.NewVisioDocument();
            var results = cmd.Invoke<IVisio.DocumentClass>();
            var doc = results.First();
            return doc;
        }

        public IVisio.PageClass New_VisioPage()
        {
            var cmd = new VisioPowerShell.Commands.NewVisioPage();
            var results = cmd.Invoke<IVisio.PageClass>();
            var page  = results.First();
            return page;
        }

        public IVisio.Shape New_VisioShape_Rectangle(double[] points)
        {
            var cmd = new VisioPowerShell.Commands.NewVisioShape();
            cmd.Type = VisioPowerShell.Commands.ShapeType.Rectangle;
            cmd.Points = points;
            var results = cmd.Invoke<IVisio.ShapeClass>();
            var shape = results.First();
            return shape;
        }

        public void Set_VisioShapeText(string s)
        {
            var cmd = new VisioPowerShell.Commands.SetVisioShapeText();
            cmd.Text = new [] {s};
            var results = cmd.Invoke();
        }

        public IVisio.ApplicationClass Get_VisioApplication()
        {
            var cmd = new VisioPowerShell.Commands.GetVisioApplication();
            var results = cmd.Invoke<IVisio.ApplicationClass>();
            var app = results.First();
            return app;
        }

        public void Close_VisioApplication()
        {
            var cmd = new VisioPowerShell.Commands.CloseVisioApplication();
            cmd.Force = true;
            var results = cmd.Invoke();
        }

    }
}