using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;
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

        public IVisio.ShapeClass Cmd_New_VisioContainer(
            string cont_master_name, 
            string cont_doc)
        {
            var doc = this.Cmd_Open_VisioDocument(cont_doc);
            var master = this.Cmd_Get_VisioMaster(cont_master_name,cont_doc);

            var cmd = new VisioPowerShell.Commands.NewVisioContainer();
            cmd.Master = master[0];
            var shape = cmd.InvokeFirst<IVisio.ShapeClass>();
            return shape ;
        }

        public List<IVisio.Shape> Cmd_New_VisioShape(
            IVisio.Master[] masters, 
            double[] points)
        {
            var cmd = new VisioPowerShell.Commands.NewVisioShape();
            cmd.Masters = masters;
            cmd.Points= points;
            var shape_list = cmd.InvokeFirst<List<IVisio.Shape>>();
            return shape_list;
        }

        public List<IVisio.Master> Cmd_Get_VisioMaster(
            string mastername, 
            string stencilname)
        {
            var doc = this.Cmd_Open_VisioDocument(stencilname);

            var cmd = new VisioPowerShell.Commands.GetVisioMaster();
            cmd.Name = mastername;
            cmd.Document = doc;
            var master = cmd.InvokeFirst<List<IVisio.Master>>();
            return master;
        }

        public List<IVisio.Master> Cmd_Get_VisioMaster(
            string mastername,
            IVisio.Document stencil)
        {
            var cmd = new VisioPowerShell.Commands.GetVisioMaster();
            cmd.Name = mastername;
            cmd.Document = stencil;
            var master = cmd.InvokeFirst<List<IVisio.Master>>();
            return master;
        }


        public IVisio.DocumentClass Cmd_Open_VisioDocument(
            string filename)
        {
            var cmd = new VisioPowerShell.Commands.OpenVisioDocument();
            cmd.Filename = filename;
            var doc = cmd.InvokeFirst<IVisio.DocumentClass>();
            return doc;
        }

        public IVisio.Document Cmd_New_VisioDocument()
        {
            var cmd = new VisioPowerShell.Commands.NewVisioDocument();
            var doc = cmd.InvokeFirst<IVisio.DocumentClass>();
            return (IVisio.Document)doc;
        }

        public System.Data.DataTable Cmd_Get_VisioShapeCells(
            IVisio.Shape[] shapes)
        {
            var cmd = new VisioPowerShell.Commands.GetVisioShapeCells();
            cmd.Shapes = shapes;
            var cells = cmd.InvokeFirst<System.Data.DataTable>();
            return cells;
        }

        public System.Data.DataTable Cmd_Get_VisioPageCells(
            IVisio.Page[] pages)
        {
            var cmd = new VisioPowerShell.Commands.GetVisioPageCells();
            cmd.Pages = pages;
            var cells = cmd.InvokeFirst<System.Data.DataTable>();
            return cells;
        }


        public IVisio.PageClass Cmd_New_VisioPage()
        {
            var cmd = new VisioPowerShell.Commands.NewVisioPage();
            var results = cmd.Invoke<IVisio.PageClass>();
            var page  = results.First();
            return page;
        }

        public IVisio.Shape Cmd_New_VisioShape(
            VisioPowerShell.Commands.ShapeType type, 
            double[] points)
        {
            var cmd = new VisioPowerShell.Commands.NewVisioShape();
            cmd.Type = type;
            cmd.Points = points;
            var shape = cmd.InvokeFirst<IVisio.ShapeClass>();
            return shape;
        }

        public void Cmd_Set_VisioText(
            string[] text, 
            IVisio.Shape[] shapes)
        {
            var cmd = new VisioPowerShell.Commands.SetVisioText();
            cmd.Text = text;
            cmd.Shapes = shapes;
            cmd.InvokeVoid();
        }

        public void Cmd_Set_VisioShapeCells(
            VisioPowerShell.Models.ShapeCells[] cells,
            IVisio.Shape[] shapes)
        {
            var cmd = new VisioPowerShell.Commands.SetVisioShapeCells();
            cmd.Cells = cells;
            cmd.Shapes = shapes;
            cmd.InvokeVoid();
        }

        public void Cmd_Set_VisioPageCells(
            VisioPowerShell.Models.PageCells[] cells,
            IVisio.Page[] pages)
        {
            var cmd = new VisioPowerShell.Commands.SetVisioPageCells();
            cmd.Cells = cells;
            cmd.Pages = pages;
            cmd.InvokeVoid();
        }

        public string[] Cmd_Get_VisioText()
        {
            var cmd = new VisioPowerShell.Commands.GetVisioText();
            var results = cmd.InvokeFirst<List<string>>();
            return results.ToArray();
        }

        public void Cmd_Close_VisioDocument(
            IVisio.Document[] documents, 
            bool force)
        {
            var cmd = new VisioPowerShell.Commands.CloseVisioDocument();
            cmd.Documents = documents;
            cmd.Force = force;
            cmd.InvokeVoid();
        }

        public IVisio.ApplicationClass Cmd_Get_VisioApplication()
        {
            var cmd = new VisioPowerShell.Commands.GetVisioApplication();
            var app = cmd.InvokeFirst<IVisio.ApplicationClass>();
            return app;
        }

        public IVisio.Page Cmd_Get_VisioPage(
            bool activepage, 
            string name)
        {
            var cmd = new VisioPowerShell.Commands.GetVisioPage();
            cmd.ActivePage = activepage;
            cmd.Name = name;
            var page = cmd.InvokeFirst<IVisio.Page>();
            return page;
        }

        public VisioPowerShell.Models.ShapeCells Cmd_New_VisioShapeCells()
        {
            var cmd = new VisioPowerShell.Commands.NewVisioShapeCells();
            var cells = cmd.InvokeFirst<VisioPowerShell.Models.ShapeCells>();
            return cells;
        }

        public VisioPowerShell.Models.PageCells Cmd_New_VisioPageCells()
        {
            var cmd = new VisioPowerShell.Commands.NewVisioPageCells();
            var cells = cmd.InvokeFirst<VisioPowerShell.Models.PageCells>();
            return cells;
        }

        public void Cmd_Close_VisioApplication()
        {
            var cmd = new VisioPowerShell.Commands.CloseVisioApplication();
            cmd.Force = true;
            cmd.InvokeVoid();
        }
    }
}