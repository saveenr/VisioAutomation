using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation_Tests
{
    [TestClass]
    public class VisioAutomationTest
    {
        private static readonly VisioApplicationSafeReference app_ref = new VisioApplicationSafeReference();
        public readonly VisioAutomation.Geometry.Size StandardPageSize = new VisioAutomation.Geometry.Size(8.5, 11);
        public readonly VisioAutomation.Geometry.Rectangle StandardPageSizeRect = new VisioAutomation.Geometry.Rectangle(new VisioAutomation.Geometry.Point(0, 0), new VisioAutomation.Geometry.Size(8.5, 11));

        public IVisio.Application GetVisioApplication()
        {
            var app = VisioAutomationTest.app_ref.GetVisioApplication();
            return app;
        }

        public IVisio.Page GetNewPage()
        {
            return this.GetNewPage(string.Empty);
        }

        public IVisio.Page GetNewPage(string suffix)
        {

            var frame = new StackFrame(2);
            var method = frame.GetMethod();
            var type = method.DeclaringType;
            var name = method.Name;
            var page = this.GetNewPage(this.StandardPageSize);
            string pagename = string.Format("{0}{1}", name,suffix);
            page.NameU = pagename;
            return page;
        }

        public IVisio.Page GetNewPage(VisioAutomation.Geometry.Size s)
        {
            var app = this.GetVisioApplication();
            var documents = app.Documents;
            if (documents.Count < 1)
            {
                var doc = documents.Add(string.Empty);
                doc.AutoRecover = false;
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
            var doc = documents.Add(string.Empty);

            var frame = new StackFrame(1);
            var method = frame.GetMethod();
            var type = method.DeclaringType;
            var name = method.Name;
            string pagename = string.Format("{0}{1}", name, "_doc");

            doc.Subject = pagename;
            return doc;

        }

        public VisioScripting.Client GetScriptingClient()
        {
            var app = this.GetVisioApplication();
            // this ensures that any debug, verbose, user , etc. messages are 
            // sent to a useful place in the unit tests
            var context = new DiagnosticDebugClientContext(); 
            var client = new VisioScripting.Client(app,context);
            return client;
        }

        public static VisioAutomation.Geometry.Size GetSize(IVisio.Shape shape)
        {
            var query = new CellQuery();
            var col_w = query.Columns.Add(VisioAutomation.ShapeSheet.SrcConstants.XFormWidth,"Width");
            var col_h = query.Columns.Add(VisioAutomation.ShapeSheet.SrcConstants.XFormHeight,"Height");

            var row = query.GetResults<double>(shape);
            double w = row[col_w];
            double h = row[col_h];
            var size = new VisioAutomation.Geometry.Size(w, h);
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
                throw new System.ArgumentNullException(nameof(window));
            }

            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            var selectargs = IVisio.VisSelectArgs.visSelect;
            window.Select(shapes, selectargs);
            var selection = window.Selection;
            var group = selection.Group();
            return group;
        }

        public static void SetPageSize(IVisio.Page page, VisioAutomation.Geometry.Size size)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var page_sheet = page.PageSheet;

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageWidth, size.Width);
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageHeight, size.Height);

            writer.Commit(page_sheet);
        }

        public static VisioAutomation.Geometry.Size GetPageSize(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var query = new CellQuery();
            var col_height = query.Columns.Add(VisioAutomation.ShapeSheet.SrcConstants.PageHeight, "PageHeight");
            var col_width = query.Columns.Add(VisioAutomation.ShapeSheet.SrcConstants.PageWidth, "PageWidth");

            var results = query.GetResults<double>(page.PageSheet);
            double height = results[col_height];
            double width = results[col_width];
            var s = new VisioAutomation.Geometry.Size(width, height);
            return s;
        }

        protected string GetTestResultsOutPath(string path)
        {
            return System.IO.Path.Combine(this.TestResultsOutFolder, path);
        }

        private static string test_result_out_folder;

        protected string TestResultsOutFolder
        {
            get
            {
                if (VisioAutomationTest.test_result_out_folder == null)
                {
                    var asm = System.Reflection.Assembly.GetExecutingAssembly();
                    VisioAutomationTest.test_result_out_folder = System.IO.Path.GetDirectoryName(asm.Location);
                }
                return VisioAutomationTest.test_result_out_folder;
            }
        }

    }
}