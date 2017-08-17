using System.Collections;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;
using VisioPowerShell.Models;
using VisioPowerShell_Tests.Framework.Extensions;

namespace VisioPowerShell_Tests
{
    public class VisioPS_Session : VisioPowerShell_Tests.Framework.PowerShellSession
    {
        private static System.Reflection.Assembly visiops_asm = typeof(VisioPowerShell.Commands.VisioCmdlet).Assembly;

        public VisioPS_Session() :
            base(visiops_asm)
        {
            
        }

        public IVisio.ShapeClass New_VisioContainer(string cont_master_name, string cont_doc)
        {
            var doc = this.Open_VisioDocument(cont_doc);
            var master = this.Get_VisioMaster(cont_master_name,cont_doc);

            var cmd = new VisioPowerShell.Commands.NewVisioContainer();
            cmd.Master = master;
            var shape = cmd.InvokeFirst<IVisio.ShapeClass>();
            return shape ;
        }

        public List<IVisio.Shape> New_VisioShape(IVisio.MasterClass master, double[] points)
        {
            var cmd = new VisioPowerShell.Commands.NewVisioShape();
            cmd.Masters = new IVisio.Master[]{ master };
            cmd.Points= points;
            var shape_list = cmd.InvokeFirst<List<IVisio.Shape>>();
            return shape_list;
        }

        public IVisio.MasterClass Get_VisioMaster(string rectangle, string basic_u_vss)
        {
            var doc = this.Open_VisioDocument(basic_u_vss);

            var cmd = new VisioPowerShell.Commands.GetVisioMaster();
            cmd.Name = rectangle;
            cmd.Document = doc;
            var master = cmd.InvokeFirst<IVisio.MasterClass>();
            return master;
        }

        public IVisio.DocumentClass Open_VisioDocument(string filename)
        {
            var cmd = new VisioPowerShell.Commands.OpenVisioDocument();
            cmd.Filename = filename;
            var doc = cmd.InvokeFirst<IVisio.DocumentClass>();
            return doc;
        }

        public IVisio.DocumentClass New_VisioDocument()
        {
            var cmd = new VisioPowerShell.Commands.NewVisioDocument();
            var doc = cmd.InvokeFirst<IVisio.DocumentClass>();
            return doc;
        }

        public System.Data.DataTable Get_VisioShapeCells(IVisio.Shape shape)
        {
            var cmd = new VisioPowerShell.Commands.GetVisioShapeSheetCells();
            cmd.Type = CellType.Page;
            cmd.Shapes = new[] {shape};
            var cells = cmd.InvokeFirst<System.Data.DataTable>();
            return cells;
        }

        public IVisio.PageClass New_VisioPage()
        {
            var cmd = new VisioPowerShell.Commands.NewVisioPage();
            var results = cmd.Invoke<IVisio.PageClass>();
            var page  = results.First();
            return page;
        }

        public IVisio.Shape New_VisioShape(VisioPowerShell.Commands.ShapeType type, double[] points)
        {
            var cmd = new VisioPowerShell.Commands.NewVisioShape();
            cmd.Type = type;
            cmd.Points = points;
            var shape = cmd.InvokeFirst<IVisio.ShapeClass>();
            return shape;
        }

        public void Set_VisioShapeText(string text, IVisio.Shape shapes)
        {
            var cmd = new VisioPowerShell.Commands.SetVisioShapeText();
            cmd.Text = new [] { text };
            cmd.Shapes = new[] {shapes};
            cmd.InvokeVoid();
        }

        public string[] Get_VisioShapeText()
        {
            var cmd = new VisioPowerShell.Commands.GetVisioShapeText();
            var results = cmd.InvokeFirst<List<string>>();
            return results.ToArray();
        }

        public void Close_VisioDocument(IVisio.Document[] docs, bool force)
        {
            var cmd = new VisioPowerShell.Commands.CloseVisioDocument();
            cmd.Documents = docs;
            cmd.Force = force;
            cmd.InvokeVoid();
        }

        public IVisio.ApplicationClass Get_VisioApplication()
        {
            var cmd = new VisioPowerShell.Commands.GetVisioApplication();
            var app = cmd.InvokeFirst<IVisio.ApplicationClass>();
            return app;
        }

        public IVisio.Page Get_VisioPage()
        {
            var cmd = new VisioPowerShell.Commands.GetVisioPage();
            var page = cmd.InvokeFirst<IVisio.Page>();
            return page;
        }

        public void Close_VisioApplication()
        {
            var cmd = new VisioPowerShell.Commands.CloseVisioApplication();
            cmd.Force = true;
            cmd.InvokeVoid();
        }
    }
}