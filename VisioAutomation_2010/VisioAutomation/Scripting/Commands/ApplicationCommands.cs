using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ApplicationCommands : CommandSet
    {
        public IVisio.Application VisioApplication { get; set; }

        public ApplicationWindowCommands Window { get; private set; }

        internal ApplicationCommands(Client client) :
            this(client, null)
        {
        }

        internal ApplicationCommands(Client client, IVisio.Application application) :
            base(client)
        {
            this.Window = new ApplicationWindowCommands(this.Client);
            this.VisioApplication = application;
        }

        public bool HasApplication
        {
            get
            {
                bool b = this.VisioApplication != null;
                this.Client.WriteVerbose("HasApplication: {0}", b);
                return b;
            }
        }

        public IVisio.Application Get()
        {
            return this.VisioApplication;
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
            var app = this.Client.Application.Get();

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
                Documents.DocumentHelper.ForceCloseAll(documents);
                app.Quit(true);
            }
            else
            {
                app.Quit();
            }
            this.VisioApplication = null;
        }

        public IVisio.Application New()
        {
            this.Client.WriteVerbose("Creating a new Instance of Visio");
            var app = new IVisio.Application();
            this.Client.WriteVerbose("Attaching that instance to current scripting client");
            this.VisioApplication = app;
            return app;
        }

        public void Undo()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.VisioApplication.Undo();
        }

        public void Redo()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.VisioApplication.Redo();
        }

        public bool Validate()
        {
            if (this.VisioApplication == null)
            {
                this.Client.WriteVerbose("Client's Application object is null");
                return false;
            }

            try
            {
                // try to do something simple, read-only, and fast with the application object
                //  if No COMException was thrown when reading ProductName property. This application instance is treated as valid

                var app_version = this.VisioApplication.ProductName;
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
                if (ApplicationCommands.visio_app_version == null)
                {
                    this.Client.Application.AssertApplicationAvailable();
                    var application = this.Client.Application.Get();
                    ApplicationCommands.visio_app_version = Application.ApplicationHelper.GetVersion(application);
                }
                return ApplicationCommands.visio_app_version;
            }            
        }

        public VA.Application.UndoScope NewUndoScope(string name)
        {
            if (this.VisioApplication == null)
            {
                throw new VisioApplicationException("Cant create UndoScope. There is no visio application attached.");
            }

            return new VA.Application.UndoScope(this.VisioApplication, name);
        }
    }
}