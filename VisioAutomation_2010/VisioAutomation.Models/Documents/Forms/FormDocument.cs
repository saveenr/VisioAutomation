

namespace VisioAutomation.Models.Documents.Forms
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

            var context = new FormRenderingContext();
            context.Application = app;
            context.Document = doc;
            context.Pages = doc.Pages;
            context.Fonts = doc.Fonts;

            this.VisioDocument = doc;

            doc.Subject = this.Subject;
            doc.Title = this.Title;
            doc.Creator = this.Creator;
            doc.Company = this.Company;

            var pages = doc.Pages;
            foreach (var formpage in this.Pages)
            {
                var page = formpage.Draw(context);
            }

            if (pages.Count > 0)
            {
                // Delete the empty first page
                var first_page = this.VisioDocument.Pages[1];
                first_page.Delete(1);
                first_page = pages[1];
                var active_window = app.ActiveWindow;
                active_window.Page = first_page;
            }
            return doc;
        }
    }
}