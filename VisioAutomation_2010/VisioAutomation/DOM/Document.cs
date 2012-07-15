using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.DOM
{
    public class Document
    {
        public PageList Pages;

        public Document()
        {
            this.Pages = new PageList();
        }

        public IVisio.Document Render(IVisio.Application app)
        {
            var appdocs = app.Documents;
            var vdoc = appdocs.Add("");
            var docpages = vdoc.Pages;
            var starpage = docpages[1];
            this.Pages.Render(starpage);
            return vdoc;
        }
    }
}