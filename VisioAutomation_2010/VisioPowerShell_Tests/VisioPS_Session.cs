using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Geometry;
using VisioPowerShell.Commands.VisioApplication;
using VisioPowerShell.Commands.VisioContainer;
using VisioPowerShell.Commands.VisioDocument;
using VisioPowerShell.Commands.VisioMaster;
using VisioPowerShell.Commands.VisioPage;
using VisioPowerShell.Commands.VisioPageCells;
using VisioPowerShell.Commands.VisioShape;
using VisioPowerShell.Commands.VisioShapeCells;
using VisioPowerShell.Commands.VisioText;
using VisioPowerShell_Tests.Framework.Extensions;
using VisioPowerShell_Tests.Framework;

namespace VisioPowerShell_Tests
{
    public class VisioPS_Session : VisioPowerShell_Tests.Framework.PowerShellSession
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

            var cmd = new NewVisioContainer();
            cmd.Master = master[0];
            var shape = cmd.InvokeFirst<IVisio.ShapeClass>();
            return shape ;
        }

        public List<IVisio.Shape> Cmd_New_VisioShape(
            IVisio.Master[] masters, 
            VisioAutomation.Geometry.Point[] points)
        {
            var cmd = new NewVisioShape();
            cmd.Master = masters;
            cmd.Points= points;
            var shape_list = cmd.Invoke<IVisio.Shape>().ToList();
            return shape_list;
        }

        public List<IVisio.Master> Cmd_Get_VisioMaster(
            string[] name, 
            string stencilname)
        {
            var doc = this.Cmd_Open_VisioDocument(stencilname);

            var cmd = new GetVisioMaster();
            cmd.Name = name;
            cmd.Document = doc;
            var master = cmd.Invoke<IVisio.Master>().ToList();
            return master;
        }

        public List<IVisio.Master> Cmd_Get_VisioMaster(
            string [] name,
            IVisio.Document stencil)
        {
            var cmd = new GetVisioMaster();
            cmd.Name = name;
            cmd.Document = stencil;
            var master = cmd.Invoke<IVisio.Master>().ToList();
            return master;
        }

        public IVisio.DocumentClass Cmd_Open_VisioDocument(
            string filename)
        {
            var cmd = new OpenVisioDocument();
            cmd.Filename = filename;
            var doc = cmd.InvokeFirst<IVisio.DocumentClass>();
            return doc;
        }

        public IVisio.Document Cmd_New_VisioDocument()
        {
            var cmd = new NewVisioDocument();
            var doc = cmd.InvokeFirst<IVisio.DocumentClass>();
            return (IVisio.Document)doc;
        }

        public System.Data.DataTable Cmd_Get_VisioShapeCells(
            IVisio.Shape[] shapes)
        {
            var cmd = new GetVisioShapeCells();
            cmd.Shape = shapes;
            var cells = cmd.InvokeFirst<System.Data.DataTable>();
            return cells;
        }

        public System.Data.DataTable Cmd_Get_VisioPageCells(
            IVisio.Page[] pages)
        {
            var cmd = new GetVisioPageCells();
            cmd.Page = pages;
            var cells = cmd.InvokeFirst<System.Data.DataTable>();
            return cells;
        }

        public IVisio.PageClass Cmd_New_VisioPage()
        {
            var cmd = new NewVisioPage();
            var results = cmd.Invoke<IVisio.PageClass>();
            var page  = results.First();
            return page;
        }

        public IVisio.Shape Cmd_New_VisioShape_rectangle(
            VisioAutomation.Geometry.Point[] points)
        {
            var cmd = new NewVisioShape();
            cmd.Rectangle = true;
            cmd.BoundingBox = new Rectangle(points[0],points[1]);
            var shape = cmd.InvokeFirst<IVisio.ShapeClass>();
            return shape;
        }

        public void Cmd_Set_VisioText(
            string[] text, 
            IVisio.Shape[] shapes)
        {
            var cmd = new SetVisioText();
            cmd.Text = text;
            cmd.Shape = shapes;
            cmd.InvokeVoid();
        }

        public void Cmd_Set_VisioShapeCells(
            VisioPowerShell.Models.ShapeCells[] cells,
            IVisio.Shape[] shapes)
        {
            var cmd = new SetVisioShapeCells();
            cmd.Cells = cells;
            cmd.Shape = shapes;
            cmd.InvokeVoid();
        }

        public void Cmd_Set_VisioPageCells(
            VisioPowerShell.Models.PageCells[] cells,
            IVisio.Page[] pages)
        {
            var cmd = new SetVisioPageCells();
            cmd.Cells = cells;
            cmd.Page = pages;
            cmd.InvokeVoid();
        }

        public string[] Cmd_Get_VisioText()
        {
            var cmd = new GetVisioText();
            var results = cmd.InvokeFirst<List<string>>();
            return results.ToArray();
        }

        public void Cmd_Close_VisioDocument(
            IVisio.Document[] documents, 
            bool force)
        {
            var cmd = new CloseVisioDocument();
            cmd.Document = documents;
            cmd.InvokeVoid();
        }

        public IVisio.ApplicationClass Cmd_Get_VisioApplication()
        {
            var cmd = new GetVisioApplication();
            var app = cmd.InvokeFirst<IVisio.ApplicationClass>();
            return app;
        }

        public IVisio.Page Cmd_Get_VisioPage(
            bool activepage, 
            string name)
        {
            var cmd = new GetVisioPage();
            cmd.ActivePage = activepage;
            cmd.Name = name;
            var page = cmd.InvokeFirst<IVisio.Page>();
            return page;
        }

        public VisioPowerShell.Models.ShapeCells Cmd_New_VisioShapeCells()
        {
            var cmd = new NewVisioShapeCells();
            var cells = cmd.InvokeFirst<VisioPowerShell.Models.ShapeCells>();
            return cells;
        }

        public VisioPowerShell.Models.PageCells Cmd_New_VisioPageCells()
        {
            var cmd = new NewVisioPageCells();
            var cells = cmd.InvokeFirst<VisioPowerShell.Models.PageCells>();
            return cells;
        }

        public void Cmd_Close_VisioApplication(bool force)
        {
            var cmd = new CloseVisioApplication();
            cmd.Force = force;
            cmd.InvokeVoid();
        }
    }
}