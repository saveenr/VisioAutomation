using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.DOM
{
    public class Document
    {
        public PageList Pages;
        private string vst ;
        private IVisio.VisMeasurementSystem? measurementSystem;
        public IVisio.Document VisioDocument;

        public Document()
        {
            this.Pages = new PageList();
        }

        public Document(string template, IVisio.VisMeasurementSystem ms) :
            this()
        {
            this.vst = template;
            this.measurementSystem = ms;
        }

        public IVisio.Document Render(IVisio.Application app)
        {
            var appdocs = app.Documents;
            IVisio.Document vdoc = null;
            if (this.vst == null)
            {
                vdoc = appdocs.Add("");
            }
            else
            {
                vdoc = appdocs.AddEx(this.vst, this.measurementSystem.Value, 0, 0);
            }
            this.VisioDocument = vdoc;
            var docpages = vdoc.Pages;
            var starpage = docpages[1];
            this.Pages.Render(starpage);
            return vdoc;
        }
    }
}