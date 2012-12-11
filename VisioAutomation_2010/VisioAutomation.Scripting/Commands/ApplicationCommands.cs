using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ApplicationCommands : CommandSet
    {
        public ApplicationWindowCommands Window { get; private set; }

        public ApplicationCommands(Session session) :
            base(session)
        {
            this.Window = new ApplicationWindowCommands(this.Session);
        }


        public void ForceClose()
        {
            var application = this.Session.VisioApplication;
            var documents = application.Documents;
            VA.Documents.DocumentHelper.ForceCloseAll(documents);
            application.Quit(true);
            this.Session.VisioApplication = null;
        }

        public IVisio.Application Attach()
        {
            var app = VA.Application.ApplicationHelper.FindRunningApplication();
            if (app == null)
            {
                throw new AutomationException("Did not find a running instance of Visio 2007");
            }

            this.Session.VisioApplication = app;

            VA.Application.ApplicationHelper.BringWindowToTop(app);

            return app;
        }

        public IVisio.Application New()
        {
            var app = new IVisio.Application();
            this.Session.VisioApplication = app;
            return app;
        }

        public void Undo()
        {
            this.Session.VisioApplication.Undo();
        }

        public void Redo()
        {
            this.Session.VisioApplication.Redo();
        }
    }
}