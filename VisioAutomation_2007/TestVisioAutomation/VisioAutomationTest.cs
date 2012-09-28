using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class VisioAutomationTest
    {
        private static VisioApplicationSafeReference app_ref = new VisioApplicationSafeReference();
        public readonly VA.Drawing.Size StandardPageSize = new VA.Drawing.Size(8.5, 11);
        public readonly VA.Drawing.Rectangle StandardPageSizeRect = new VA.Drawing.Rectangle(new VA.Drawing.Point(0, 0), new VA.Drawing.Size(8.5, 11));

        public IVisio.Application GetVisioApplication()
        {
            var app = app_ref.GetVisioApplication();
            return app;
        }

        public IVisio.Page GetNewPage()
        {
            return this.GetNewPage("");
        }

        public IVisio.Page GetNewPage(string suffix)
        {

            var frame = new System.Diagnostics.StackFrame(2);
            var method = frame.GetMethod();
            var type = method.DeclaringType;
            var name = method.Name;
            var page = GetNewPage(StandardPageSize);
            string pagename = string.Format("{0}{1}", name,suffix);
            page.NameU = pagename;
            return page;
        }

        public IVisio.Page GetNewPage(VA.Drawing.Size s)
        {
            var app = GetVisioApplication();
            var documents = app.Documents;
            if (documents.Count < 1)
            {
                documents.Add(string.Empty);
            }
            var active_document = app.ActiveDocument;
            var pages = active_document.Pages;
            var page = pages.Add();
            page.Background = 0;
            page.SetSize(s);

            return page;
        }

        public IVisio.Document GetNewDoc()
        {
            var app = GetVisioApplication();
            var documents = app.Documents;
            var doc = documents.Add("");

            var frame = new System.Diagnostics.StackFrame(1);
            var method = frame.GetMethod();
            var type = method.DeclaringType;
            var name = method.Name;
            string pagename = string.Format("{0}{1}", name, "_doc");

            doc.Subject = pagename;
            return doc;

        }

        public VA.Scripting.Session GetScriptingSession()
        {
            var app = GetVisioApplication();
            var scriptingsession = new VA.Scripting.Session(app);
            return scriptingsession;
        }

        public static VA.Drawing.Size GetSize(IVisio.Shape shape)
        {
            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_w = query.AddColumn(VA.ShapeSheet.SRCConstants.Width);
            var col_h = query.AddColumn(VA.ShapeSheet.SRCConstants.Height);

            var table = query.GetResults<double>(shape);
            double w = table[0, col_w];
            double h = table[0, col_h];
            var size = new VA.Drawing.Size(w, h);
            return size;
        }

        public static void ResetDoc(IVisio.Document doc)
        {
            var pages = doc.Pages;

            var target_pages = new System.Collections.Generic.List<IVisio.Page>(pages.Count);
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
    }
}