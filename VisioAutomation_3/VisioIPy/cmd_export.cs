using VisioAutomation.Scripting;
using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public void ExportSelection(string filename)
        {
            this.ScriptingSession.Export.ExportSelectionToFile(filename);
        }

        public void ExportPage(string filename)
        {
            this.ScriptingSession.Export.ExportPageToFile(filename);
        }

        public void ExportPages(string filename)
        {
            this.ScriptingSession.Export.ExportPagesToFiles(filename);
        }

        public void ExportSelectionAsSVGXHTML(string filename)
        {
            var ss = this.ScriptingSession;
            ss.Export.ExportSelectionAsSVGXHTML(filename);
        }

        public void ExportSelectionAsXAML(string filename)
        {
            var ss = this.ScriptingSession;
            VAS.XamlTune.XamlTuneHelper.ExportSelectionAsXAML(ss, filename);
        }
    }
}