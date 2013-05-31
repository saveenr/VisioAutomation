using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.DOM
{
    public class Document
    {
        public PageList Pages;
        private readonly string vst_template_file ;
        private IVisio.VisMeasurementSystem? measurementSystem;
        public IVisio.Document VisioDocument;

        public Document()
        {
            this.Pages = new PageList();
        }

        public Document(string template, IVisio.VisMeasurementSystem ms) :
            this()
        {
            this.vst_template_file = template;
            this.measurementSystem = ms;
        }

        public IVisio.Document Render(IVisio.Application app)
        {
            var appdocs = app.Documents;
            IVisio.Document vdoc = null;
            if (this.vst_template_file == null)
            {
                vdoc = appdocs.Add("");
            }
            else
            {
                const int flags = 0; // (int)IVisio.VisOpenSaveArgs.visAddDocked;
                const int langid = 0;
                vdoc = appdocs.AddEx(this.vst_template_file, this.measurementSystem.Value, flags, langid);
            }
            this.VisioDocument = vdoc;
            var docpages = vdoc.Pages;
            var startpage = docpages[1];
            this.Pages.Render(startpage);
            return vdoc;
        }
    }
}