using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class VisioAutomationTest
    {
        private static readonly VisioApplicationSafeReference app_ref = new VisioApplicationSafeReference();
        public readonly VisioAutomation.Drawing.Size StandardPageSize = new VisioAutomation.Drawing.Size(8.5, 11);
        public readonly VisioAutomation.Drawing.Rectangle StandardPageSizeRect = new VisioAutomation.Drawing.Rectangle(new VisioAutomation.Drawing.Point(0, 0), new VisioAutomation.Drawing.Size(8.5, 11));

        public IVisio.Application GetVisioApplication()
        {
            var app = VisioAutomationTest.app_ref.GetVisioApplication();
            return app;
        }

        public IVisio.Page GetNewPage()
        {
            return this.GetNewPage("");
        }

        public IVisio.Page GetNewPage(string suffix)
        {

            var frame = new StackFrame(2);
            var method = frame.GetMethod();
            var type = method.DeclaringType;
            var name = method.Name;
            var page = this.GetNewPage(this.StandardPageSize);
            string pagename = $"{name}{suffix}";
            page.NameU = pagename;
            return page;
        }

        public IVisio.Page GetNewPage(VisioAutomation.Drawing.Size s)
        {
            var app = this.GetVisioApplication();
            var documents = app.Documents;
            if (documents.Count < 1)
            {
                documents.Add(string.Empty);
            }
            var active_document = app.ActiveDocument;
            var pages = active_document.Pages;
            var page = pages.Add();
            page.Background = 0;
            VisioAutomationTest.SetPageSize(page, s);

            return page;
        }

        public IVisio.Document GetNewDoc()
        {
            var app = this.GetVisioApplication();
            var documents = app.Documents;
            var doc = documents.Add("");

            var frame = new StackFrame(1);
            var method = frame.GetMethod();
            var type = method.DeclaringType;
            var name = method.Name;
            string pagename = $"{name}{"_doc"}";

            doc.Subject = pagename;
            return doc;

        }

        public VisioAutomation.Scripting.Client GetScriptingClient()
        {
            var app = this.GetVisioApplication();
            // this ensures that any debug, verbose, user , etc. messages are 
            // sent to a useful place in the unit tests
            var context = new DiagnosticDebugContext(); 
            var client = new VisioAutomation.Scripting.Client(app,context);
            return client;
        }

        public static VisioAutomation.Drawing.Size GetSize(IVisio.Shape shape)
        {
            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();
            var col_w = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Width,"Width");
            var col_h = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Height,"Height");

            var table = query.GetResults<double>(shape);
            double w = table[col_w];
            double h = table[col_h];
            var size = new VisioAutomation.Drawing.Size(w, h);
            return size;
        }

        public static void ResetDoc(IVisio.Document doc)
        {
            var pages = doc.Pages;

            var target_pages = new List<IVisio.Page>(pages.Count);
            foreach (IVisio.Page p in pages)
            {
                target_pages.Add(p);
            }

            var empty_page = pages.Add();

            foreach (IVisio.Page p in target_pages)
            {
                p.Delete(1);
            }
        }

        /// Selects a series of shapes and groups them into one shape
        public static IVisio.Shape SelectAndGroup(IVisio.Window window, IEnumerable<IVisio.Shape> shapes)
        {
            if (window == null)
            {
                throw new System.ArgumentNullException("window");
            }

            if (shapes == null)
            {
                throw new System.ArgumentNullException("shapes");
            }

            var selectargs = IVisio.VisSelectArgs.visSelect;
            window.Select(shapes, selectargs);
            var selection = window.Selection;
            var group = selection.Group();
            return group;
        }

        public static void SetPageSize(IVisio.Page page, VisioAutomation.Drawing.Size size)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var page_sheet = page.PageSheet;

            var update = new VisioAutomation.ShapeSheet.Update(2);
            update.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.PageWidth, size.Width);
            update.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.PageHeight, size.Height);
            update.Execute(page_sheet);
        }

        public static VisioAutomation.Drawing.Size GetPageSize(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();
            var col_height = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.PageHeight, "PageHeight");
            var col_width = query.AddCell(VisioAutomation.ShapeSheet.SRCConstants.PageWidth, "PageWidth");
            var results = query.GetResults<double>(page.PageSheet);
            double height = results[col_height];
            double width = results[col_width];
            var s = new VisioAutomation.Drawing.Size(width, height);
            return s;
        }

        protected string GetTestResultsOutPath(string path)
        {
            return System.IO.Path.Combine(this.TestResultsOutFolder, path);
        }

        private static string tr_out_folder;

        protected string TestResultsOutFolder
        {
            get
            {
                if (VisioAutomationTest.tr_out_folder == null)
                {
                    var asm = System.Reflection.Assembly.GetExecutingAssembly();
                    VisioAutomationTest.tr_out_folder = System.IO.Path.GetDirectoryName(asm.Location);
                }
                return VisioAutomationTest.tr_out_folder;
            }
        }

    }
}