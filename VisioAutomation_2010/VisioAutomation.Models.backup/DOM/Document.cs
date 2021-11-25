using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Dom
{
    public class Document
    {
        public PageList Pages;
        public IVisio.Document VisioDocument;

        private readonly string _vst_template_file ;
        private readonly IVisio.VisMeasurementSystem _measurement_system;

        public Document()
        {
            this.Pages = new PageList();
            this._measurement_system = IVisio.VisMeasurementSystem.visMSDefault;
        }

        public Document(string template, IVisio.VisMeasurementSystem ms) :
            this()
        {
            this._vst_template_file = template;
            this._measurement_system = ms;
        }

        public IVisio.Document Render(IVisio.Application app)
        {
            var appdocs = app.Documents;
            IVisio.Document doc = null;
            if (this._vst_template_file == null)
            {
                doc = appdocs.Add(string.Empty);
            }
            else
            {
                const int flags = 0; // (int)IVisio.VisOpenSaveArgs.visAddDocked;
                const int langid = 0;
                doc = appdocs.AddEx(this._vst_template_file, this._measurement_system, flags, langid);
            }
            this.VisioDocument = doc;
            var docpages = doc.Pages;
            var startpage = docpages[1];
            this.Pages.Render(startpage);
            return doc;
        }
    }
}