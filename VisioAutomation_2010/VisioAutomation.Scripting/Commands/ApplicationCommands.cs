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
            this.CheckVisioApplicationAvailable();

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
                throw new VA.Scripting.VisioApplicationException("Did not find a running instance of Visio 2010 or above");
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
            this.CheckVisioApplicationAvailable();
            this.Session.VisioApplication.Undo();
        }

        public void Redo()
        {
            this.CheckVisioApplicationAvailable();
            this.Session.VisioApplication.Redo();
        }

        public bool Validate()
        {
            var app = this.Session.VisioApplication;

            if (app == null)
            {
                this.Session.WriteVerbose("Session's Application object is null");
                return false;
            }
            else
            {
                this.Session.WriteVerbose("Session's Application object is not null");
                try
                {
                    this.Session.WriteVerbose("Attempting to read Visio Application's Version property");
                    // try to do something simple, read-only, and fast with the application object
                    var app_version = app.Version;
                    this.Session.WriteVerbose(
                        "No COMException was thrown when reading Version property. This application instance seems valid");
                    return true;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    this.Session.WriteVerbose("COMException thrown");
                    this.Session.WriteVerbose("This application instance is invalid");
                    // If a COMException is thrown, this indicates that the
                    // application object is invalid
                    return false;
                }
                catch (System.Exception exc)
                {
                    this.Session.WriteVerbose("An exception besides COMException was thrown");
                    // just re-raise it.
                    throw;
                }
            }
        }
    }
}