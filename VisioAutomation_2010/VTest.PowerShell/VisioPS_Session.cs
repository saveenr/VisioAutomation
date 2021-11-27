using System.Collections.Generic;
using System.Linq;
using VTest.PowerShell.Framework;
using VTest.PowerShell.Framework.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VTest.PowerShell
{
    public class VisioPS_Session : PowerShellSession
    {
        public VisioPS_Session()
        {
            // Find the path to the assembly
            var visiops_asm = typeof(VisioPowerShell.Commands.VisioCmdlet).Assembly;
            var modules = new[] { visiops_asm.Location };
            this._sessionstate.ImportPSModule(modules);
        }

        public IVisio.ShapeClass Cmd_New_VisioContainer(
            string cont_master_name, 
            string cont_doc)
        {
            var doc = this.Cmd_Open_VisioDocument(cont_doc);
            var master = this.Cmd_Get_VisioMaster(PsArray.From(cont_master_name),cont_doc);

            var cmd = new VisioPowerShell.Commands.VisioContainer.NewVisioContainer();
            cmd.Master = master[0];
            var shape = cmd.InvokeFirst<IVisio.ShapeClass>();
            return shape ;
        }

        public List<IVisio.Shape> Cmd_New_VisioShape(
            IVisio.Master[] masters, 
            VisioAutomation.Core.Point[] points)
        {
            var cmd = new VisioPowerShell.Commands.VisioShape.NewVisioShape();
            cmd.Master = masters;
            cmd.Position = points;
            var shape_list = cmd.Invoke<IVisio.Shape>().ToList();
            return shape_list;
        }

        public List<IVisio.Master> Cmd_Get_VisioMaster(
            string[] name, 
            string stencilname)
        {
            var doc = this.Cmd_Open_VisioDocument(stencilname);

            var cmd = new VisioPowerShell.Commands.VisioMaster.GetVisioMaster();
            cmd.Name = name;
            cmd.Document = doc;
            var master = cmd.Invoke<IVisio.Master>().ToList();
            return master;
        }

        public List<IVisio.Master> Cmd_Get_VisioMaster(
            string [] name,
            IVisio.Document stencil)
        {
            var cmd = new VisioPowerShell.Commands.VisioMaster.GetVisioMaster();
            cmd.Name = name;
            cmd.Document = stencil;
            var master = cmd.Invoke<IVisio.Master>().ToList();
            return master;
        }

        public IVisio.DocumentClass Cmd_Open_VisioDocument(
            string filename)
        {
            var cmd = new VisioPowerShell.Commands.VisioDocument.OpenVisioDocument();
            cmd.Filename = filename;
            var doc = cmd.InvokeFirst<IVisio.DocumentClass>();
            return doc;
        }

        public IVisio.Document Cmd_New_VisioDocument()
        {
            var cmd = new VisioPowerShell.Commands.VisioDocument.NewVisioDocument();
            var doc = cmd.InvokeFirst<IVisio.DocumentClass>();
            return (IVisio.Document)doc;
        }

        public System.Data.DataTable Cmd_Get_VisioShapeCells(
            IVisio.Shape[] shapes)
        {
            var cmd = new VisioPowerShell.Commands.VisioShapeCells.GetVisioShapeCells();
            cmd.Shape = shapes;
            var cells = cmd.InvokeFirst<System.Data.DataTable>();
            return cells;
        }

        public System.Data.DataTable Cmd_Get_VisioPageCells(
            IVisio.Page[] pages)
        {
            var cmd = new VisioPowerShell.Commands.VisioPageCells.GetVisioPageCells();
            cmd.Page = pages;
            var cells = cmd.InvokeFirst<System.Data.DataTable>();
            return cells;
        }

        public IVisio.PageClass Cmd_New_VisioPage()
        {
            var cmd = new VisioPowerShell.Commands.VisioPage.NewVisioPage();
            var results = cmd.Invoke<IVisio.PageClass>();
            var page  = results.First();
            return page;
        }

        public IVisio.Shape Cmd_New_VisioShape_rectangle(
            VisioAutomation.Core.Point[] points)
        {
            var cmd = new VisioPowerShell.Commands.VisioShape.NewVisioShape();
            cmd.Rectangle = true;
            cmd.BoundingBox = new VisioAutomation.Core.Rectangle(points[0],points[1]);
            var shape = cmd.InvokeFirst<IVisio.ShapeClass>();
            return shape;
        }

        public void Cmd_Set_VisioText(
            string[] text, 
            IVisio.Shape[] shapes)
        {
            var cmd = new VisioPowerShell.Commands.VisioText.SetVisioText();
            cmd.Text = text;
            cmd.Shape = shapes;
            cmd.InvokeVoid();
        }

        public void Cmd_Set_VisioShapeCells(
            VisioPowerShell.Models.ShapeCells[] cells,
            IVisio.Shape[] shapes)
        {
            var cmd = new VisioPowerShell.Commands.VisioShapeCells.SetVisioShapeCells();
            cmd.Cells = cells;
            cmd.Shape = shapes;
            cmd.InvokeVoid();
        }

        public void Cmd_Set_VisioPageCells(
            VisioPowerShell.Models.PageCells[] cells,
            IVisio.Page[] pages)
        {
            var cmd = new VisioPowerShell.Commands.VisioPageCells.SetVisioPageCells();
            cmd.Cells = cells;
            cmd.Page = pages;
            cmd.InvokeVoid();
        }

        public string[] Cmd_Get_VisioText()
        {
            var cmd = new VisioPowerShell.Commands.VisioText.GetVisioText();
            var results = cmd.InvokeFirst<List<string>>();
            return results.ToArray();
        }

        public void Cmd_Close_VisioDocument(
            IVisio.Document[] documents, 
            bool force)
        {
            var cmd = new VisioPowerShell.Commands.VisioDocument.CloseVisioDocument();
            cmd.Document = documents;
            cmd.InvokeVoid();
        }

        public IVisio.ApplicationClass Cmd_Get_VisioApplication()
        {
            var cmd = new VisioPowerShell.Commands.VisioApplication.GetVisioApplication();
            var app = cmd.InvokeFirst<IVisio.ApplicationClass>();
            return app;
        }

        public IVisio.Page Cmd_Get_VisioPage(
            bool activepage, 
            string name)
        {
            var cmd = new VisioPowerShell.Commands.VisioPage.GetVisioPage();
            cmd.ActivePage = activepage;
            cmd.Name = new [] {name};
            var page = cmd.InvokeFirst<IVisio.Page>();
            return page;
        }

        public VisioPowerShell.Models.ShapeCells Cmd_New_VisioShapeCells()
        {
            var cmd = new VisioPowerShell.Commands.VisioShapeCells.NewVisioShapeCells();
            var cells = cmd.InvokeFirst<VisioPowerShell.Models.ShapeCells>();
            return cells;
        }

        public VisioPowerShell.Models.PageCells Cmd_New_VisioPageCells()
        {
            var cmd = new VisioPowerShell.Commands.VisioPageCells.NewVisioPageCells();
            var cells = cmd.InvokeFirst<VisioPowerShell.Models.PageCells>();
            return cells;
        }

        public void Cmd_Close_VisioApplication(bool force)
        {
            var cmd = new VisioPowerShell.Commands.VisioApplication.CloseVisioApplication();
            cmd.InvokeVoid();
        }
    }
}