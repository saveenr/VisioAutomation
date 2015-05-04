using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Drawing;
using VisioAutomation.Extensions;
using VisioAutomation.Scripting;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class VisioAutomationTest
    {
        private static readonly VisioApplicationSafeReference app_ref = new VisioApplicationSafeReference();
        public readonly Size StandardPageSize = new Size(8.5, 11);
        public readonly Rectangle StandardPageSizeRect = new Rectangle(new Point(0, 0), new Size(8.5, 11));

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
            string pagename = string.Format("{0}{1}", name,suffix);
            page.NameU = pagename;
            return page;
        }

        public IVisio.Page GetNewPage(Size s)
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
            string pagename = string.Format("{0}{1}", name, "_doc");

            doc.Subject = pagename;
            return doc;

        }

        public Client GetScriptingClient()
        {
            var app = this.GetVisioApplication();
            // this ensures that any debug, verbose, user , etc. messages are 
            // sent to a useful place in the unit tests
            var context = new DiagnosticDebugContext(); 
            var client = new Client(app,context);
            return client;
        }

        public static Size GetSize(IVisio.Shape shape)
        {
            var query = new CellQuery();
            var col_w = query.AddCell(SRCConstants.Width,"Width");
            var col_h = query.AddCell(SRCConstants.Height,"Height");

            var table = query.GetResults<double>(shape);
            double w = table[col_w];
            double h = table[col_h];
            var size = new Size(w, h);
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
                throw new ArgumentNullException("window");
            }

            if (shapes == null)
            {
                throw new ArgumentNullException("shapes");
            }

            var selectargs = IVisio.VisSelectArgs.visSelect;
            window.Select(shapes, selectargs);
            var selection = window.Selection;
            var group = selection.Group();
            return group;
        }

        public static void SetPageSize(IVisio.Page page, Size size)
        {
            if (page == null)
            {
                throw new ArgumentNullException("page");
            }

            var page_sheet = page.PageSheet;

            var update = new Update(2);
            update.SetFormula(SRCConstants.PageWidth, size.Width);
            update.SetFormula(SRCConstants.PageHeight, size.Height);
            update.Execute(page_sheet);
        }

        public static Size GetPageSize(IVisio.Page page)
        {
            if (page == null)
            {
                throw new ArgumentNullException("page");
            }

            var query = new CellQuery();
            var col_height = query.AddCell(SRCConstants.PageHeight, "PageHeight");
            var col_width = query.AddCell(SRCConstants.PageWidth, "PageWidth");
            var results = query.GetResults<double>(page.PageSheet);
            double height = results[col_height];
            double width = results[col_width];
            var s = new Size(width, height);
            return s;
        }

        protected string GetTestResultsOutPath(string path)
        {
            return Path.Combine(this.TestResultsOutFolder, path);
        }

        private static string tr_out_folder;

        protected string TestResultsOutFolder
        {
            get
            {
                if (VisioAutomationTest.tr_out_folder == null)
                {
                    var asm = Assembly.GetExecutingAssembly();
                    VisioAutomationTest.tr_out_folder = Path.GetDirectoryName(asm.Location);
                }
                return VisioAutomationTest.tr_out_folder;
            }
        }

    }
}