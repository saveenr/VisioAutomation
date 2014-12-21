using System.Collections.Generic;
using VisioAutomation.Drawing;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Forms
{
    public class FormDocument
    {
        public string Subject ;
        public string Title ;
        public string Creator ;
        public string Company;
        public List<FormPage> Pages;
        public IVisio.Document VisioDocument;

        public FormDocument()
        {
            this.Pages = new List<FormPage>();
        }

        public IVisio.Document Render(IVisio.Application app)
        {

            var docs = app.Documents;
            var doc = docs.Add("");

            var ctx = new FormRenderingContext(app);
            ctx.Application = app;
            ctx.Document = doc;
            ctx.Pages = doc.Pages;
            ctx.Fonts = doc.Fonts;

            this.VisioDocument = doc;

            doc.Subject = this.Subject;
            doc.Title = this.Title;
            doc.Creator = this.Creator;
            doc.Company = this.Company;

            var pages = doc.Pages;
            foreach (var formpage in this.Pages)
            {
                var page = formpage.Draw(ctx);
            }

            if (pages.Count > 0)
            {
                // Delete the empty first page
                var first_page = VisioDocument.Pages[1];
                first_page.Delete(1);
                first_page = pages[1];
                var active_window = app.ActiveWindow;
                active_window.Page = first_page;
            }
            return doc;
        }
    }
}