using VisioAutomation;
using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IVisio.Page ActivePage
        {
            get { return this.ScriptingSession.Page.GetPage(); }
        }

        public void ResizePageToFitContents()
        {
            this.ScriptingSession.Page.ResizeToFitContents(new VA.Drawing.Size(0, 0), true);
        }

        public void ResizePageToFitContents(double w, double h)
        {
            this.ScriptingSession.Page.ResizeToFitContents(new VA.Drawing.Size(w, h), true);
        }

        public VA.Drawing.Size GetPageSize()
        {
            return this.ScriptingSession.Page.GetPageSize();
        }

        public void SetPageSize(double w, double h)
        {
            this.SetPageSize(new VA.Drawing.Size(w, h));
        }

        public void SetPageSize(VA.Drawing.Size size)
        {
            this.ScriptingSession.Page.SetPageSize(size);
        }

        public void DuplicatePage()
        {
            this.ScriptingSession.Page.DuplicatePage();
        }

        public void DuplicatePageToNewDrawing()
        {
            this.ScriptingSession.Page.DuplicatePageToNewDocument();
        }

        public IVisio.Page NewPage()
        {
            return this.ScriptingSession.Page.NewPage(null, false);
        }

        public IVisio.Page NewPage(double w, double h)
        {
            var size = new VA.Drawing.Size(w, h);
            return this.ScriptingSession.Page.NewPage(size, false);
        }

        public IVisio.Page NewBackgroundPage()
        {
            return this.ScriptingSession.Page.NewPage(null, true);
        }

        public IVisio.Page NewBackgroundPage(double w, double h)
        {
            var size = new VA.Drawing.Size(w, h);
            return this.ScriptingSession.Page.NewPage(size, true);
        }

        public void SetBackgroundPage(string name)
        {
            this.ScriptingSession.Page.SetBackgroundPage(name);
        }

        public void GotoPage(PageNavigation pagenav)
        {
            this.ScriptingSession.Page.NavigateToPage(pagenav);
        }

        public void ResetPageOrigin()
        {
            this.ScriptingSession.Page.ResetPageOrigin();
        }
    }
}