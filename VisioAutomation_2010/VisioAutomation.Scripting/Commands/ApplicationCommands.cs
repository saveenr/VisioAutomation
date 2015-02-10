using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ApplicationCommands : CommandSet
    {
        public ApplicationWindowCommands Window { get; private set; }

        public ApplicationCommands(Client client) :
            base(client)
        {
            this.Window = new ApplicationWindowCommands(this.Client);
        }

        public bool HasApplication
        {
            get
            {
                bool b = this.Client.VisioApplication != null;
                this.Client.WriteVerbose("HasApplication: {0}", b);
                return b;
            }
        }

        public void AssertApplicationAvailable()
        {
            var has_app = this.Client.Application.HasApplication;
            if (!has_app)
            {
                throw new VisioApplicationException("No Visio Application available");
            }
        }

        public void Close(bool force)
        {
            var app = this.Client.VisioApplication;

            if (app == null)
            {
                this.Client.WriteWarning("There is no Visio Application to stop");
                return;
            }

            if (force)
            {
                // If you want to force the thing to close
                // it will require closing all documents and then quiting
                var documents = app.Documents;
                VA.Documents.DocumentHelper.ForceCloseAll(documents);
                app.Quit(true);
            }
            else
            {
                app.Quit();
            }
            this.Client.VisioApplication = null;
        }

        public IVisio.Application FindRunning()
        {
            if (VisioAutomation.Scripting.UACHelper.IsUacEnabled)
            {
                this.Client.WriteVerbose("UAC Enabled");
            }

            if (VisioAutomation.Scripting.UACHelper.IsProcessElevated)
            {
                this.Client.WriteVerbose("Running in Elevated Process");
                this.Client.WriteWarning("Having an Elevated Process with UAC Enabled will cause Running Applications to not be found");
            }

            var app = VA.Application.ApplicationHelper.FindRunningApplication();
            return app;
        }


        public IVisio.Application Attach()
        {
            if (this.Client.VisioApplication != null)
            {
                this.Client.WriteWarning("Already connected to an instance");
            }

            var app = this.FindRunning();
            if (app == null)
            {
                throw new VA.Scripting.VisioApplicationException("Did not find a running instance of Visio 2010 or above");
            }

            this.Client.WriteVerbose("Attaching to an instance");

            this.Client.VisioApplication = app;

            VA.Application.ApplicationHelper.BringWindowToTop(app);

            return app;
        }

        public IVisio.Application New()
        {
            this.Client.WriteVerbose("Creating a new Instance of Visio");
            var app = new IVisio.Application();
            this.Client.WriteVerbose("Attaching that instance to current scripting client");
            this.Client.VisioApplication = app;
            return app;
        }

        public IVisio.Application SafeNew()
        {
            if (this.Client.Application.HasApplication == false)
            {
                // no app - let's create one
                return this.Client.Application.New();
            }

            if (this.Client.Application.Validate())
            {
                // the app instance is OK
                return this.Client.VisioApplication;
            }
            else
            {
                // it's isn't good. Get a new one
                this.Client.Application.New();
                return this.Client.VisioApplication;
            }
        }

        public void Undo()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.VisioApplication.Undo();
        }

        public void Redo()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.VisioApplication.Redo();
        }

        public bool Validate()
        {
            var app = this.Client.VisioApplication;

            if (app == null)
            {
                this.Client.WriteVerbose("Client's Application object is null");
                return false;
            }

            try
            {
                // try to do something simple, read-only, and fast with the application object
                //  if No COMException was thrown when reading ProductName property. This application instance is treated as valid

                var app_version = app.ProductName;
                this.Client.WriteVerbose("Application validated");
                return true;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                this.Client.WriteVerbose("COMException thrown during validation. Treating as invalid application");
                // If a COMException is thrown, this indicates that the
                // application object is invalid
                return false;
            }
        }

        private static System.Version visio_app_version;

        public System.Version Version
        {
            get
            {
                if (visio_app_version == null)
                {
                    this.Client.Application.AssertApplicationAvailable();
                    visio_app_version = VA.Application.ApplicationHelper.GetVersion(this.Client.VisioApplication);
                }
                return visio_app_version;
            }            
        }
    }
}