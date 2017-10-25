using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class ApplicationCommands : CommandSet
    {
        private static System.Version visio_app_version;

        private IVisio.Application _active_application;
        
        internal ApplicationCommands(Client client) :
            this(client, null)
        {
        }

        internal ApplicationCommands(Client client, IVisio.Application application) :
            base(client)
        {
            this._active_application = application;
        }

        public bool HasActiveApplication
        {
            get
            {
                bool b = this._active_application != null;
                this._client.Output.WriteVerbose("HasApplication: {0}", b);
                return b;
            }
        }

        public IVisio.Application GetActiveApplication()
        {
            return this._active_application;
        }

        public void SetActiveApplication(IVisio.Application app)
        {
            this._active_application = app;
        }

        public void AssertHasActiveApplication()
        {
            var has_app = this._client.Application.HasActiveApplication;
            if (!has_app)
            {
                throw new System.ArgumentException("No Visio Application available");
            }
        }

        public void CloseActiveApplication(bool force)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            var app = cmdtarget.Application;

            if (app == null)
            {
                this._client.Output.WriteWarning("There is no Visio Application to stop");
                return;
            }

            if (force)
            {
                // If you want to force the thing to close
                // it will require closing all documents and then quiting
                var documents = app.Documents;

                while (documents.Count > 0)
                {
                    var active_document = app.ActiveDocument;
                    active_document.Close(true);
                }

                app.Quit(true);
            }
            else
            {
                app.Quit();
            }
            this._active_application = null;
        }

        public IVisio.Application NewActiveApplication()
        {
            this._client.Output.WriteVerbose("Creating a new Instance of Visio");
            var app = new IVisio.Application();
            this._client.Output.WriteVerbose("Attaching that instance to current scripting client");
            this._active_application = app;
            return app;
        }

        public bool ValidateActiveApplication()
        {
            if (this._active_application == null)
            {
                this._client.Output.WriteVerbose("Client's Application object is null");
                return false;
            }

            try
            {
                // try to do something simple, read-only, and fast with the application object
                //  if No COMException was thrown when reading ProductName property. This application instance is treated as valid

                var app_version = this._active_application.ProductName;
                this._client.Output.WriteVerbose("Application validated");
                return true;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                this._client.Output.WriteVerbose("COMException thrown during validation. Treating as invalid application");
                // If a COMException is thrown, this indicates that the
                // application object is invalid
                return false;
            }
        }

        public System.Version ApplicationVersion
        {
            get
            {
                if (ApplicationCommands.visio_app_version == null)
                {
                    var cmdtarget = this._client.GetCommandTargetApplication();

                    var application = cmdtarget.Application;
                    ApplicationCommands.visio_app_version = VisioAutomation.Application.ApplicationHelper.GetVersion(application);
                }
                return ApplicationCommands.visio_app_version;
            }            
        }
    }
}