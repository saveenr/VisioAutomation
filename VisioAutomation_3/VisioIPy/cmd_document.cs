using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IVisio.Document NewDrawing()
        {
            var doc = this.ScriptingSession.Document.NewDocument();
            return doc;
        }

        public IVisio.Document NewStencil()
        {
            var doc = this.ScriptingSession.Document.NewStencil();
            return doc;
        }

        public void CloseAllWithoutSaving()
        {
            this.ScriptingSession.Document.CloseAllDocumentsWithoutSaving();
        }

        public IVisio.Document Open(string filename)
        {
            var doc = this.Application.Documents.Open(filename);
            return doc;
        }

        public IVisio.Document OpenStencil(string filename)
        {
            var doc = this.ScriptingSession.Document.OpenStencil(filename);
            return doc;
        }

        public void Save()
        {
            this.ScriptingSession.Document.SaveDocument();
        }

        public void SaveAs(string filename)
        {
            this.ScriptingSession.Document.SaveDocumentAs(filename);
        }

        public void CloseWithoutSaving()
        {
            this.ScriptingSession.Document.CloseDocument(true);
        }

        public bool HasActiveDrawing()
        {
            var ss = this.ScriptingSession;
            return ss.HasActiveDrawing();
        }
    }
}