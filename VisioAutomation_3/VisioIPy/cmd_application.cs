using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;

using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IVisio.Application Application
        {
            get { return this.m_app; }
            private set { this.m_app = value; }
        }

        public void Attach()
        {
            detach();
            this.Application = VAS.Session.AttachToRunningApplication();
            print_app_window_text();
        }

        private void detach()
        {
            if (this.Application != null)
            {
                this.WriteLine("Unbinding from an instance of Visio 2007");
                this.Application = null;
            }
        }

        public void Attach(IVisio.Application app)
        {
            if (app == null)
            {
                throw new System.ArgumentNullException("app");
            }

            this.detach();
            this.Application = app;
            print_app_window_text();
        }

        public void NewApplication()
        {
            var scriptingsession = this.ScriptingSession;
            this.Application = scriptingsession.StartNewApplication();
        }

        public void Undo()
        {
            this.ScriptingSession.Undo();
        }

        public void Redo()
        {
            this.ScriptingSession.Redo();
        }
    }
}