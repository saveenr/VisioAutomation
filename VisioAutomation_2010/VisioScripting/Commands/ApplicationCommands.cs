using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioScripting.Commands
{
    public class ApplicationCommands : CommandSet
    {
        private static System.Version visio_app_version;

        public IVisio.Application VisioApplication { get; set; }

        internal ApplicationCommands(Client client) :
            this(client, null)
        {
        }

        internal ApplicationCommands(Client client, IVisio.Application application) :
            base(client)
        {
            this.VisioApplication = application;
        }

        public bool HasApplication
        {
            get
            {
                bool b = this.VisioApplication != null;
                this._client.Output.WriteVerbose("HasApplication: {0}", b);
                return b;
            }
        }

        public IVisio.Application Get()
        {
            return this.VisioApplication;
        }

        public void AssertApplicationAvailable()
        {
            var has_app = this._client.Application.HasApplication;
            if (!has_app)
            {
                throw new System.ArgumentException("No Visio Application available");
            }
        }

        public void Close(bool force)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);

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
            this.VisioApplication = null;
        }

        public IVisio.Application New()
        {
            this._client.Output.WriteVerbose("Creating a new Instance of Visio");
            var app = new IVisio.Application();
            this._client.Output.WriteVerbose("Attaching that instance to current scripting client");
            this.VisioApplication = app;
            return app;
        }

        public void Undo()
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);

            this.VisioApplication.Undo();
        }

        public void Redo()
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);

            this.VisioApplication.Redo();
        }

        public bool Validate()
        {
            if (this.VisioApplication == null)
            {
                this._client.Output.WriteVerbose("Client's Application object is null");
                return false;
            }

            try
            {
                // try to do something simple, read-only, and fast with the application object
                //  if No COMException was thrown when reading ProductName property. This application instance is treated as valid

                var app_version = this.VisioApplication.ProductName;
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

        public System.Version Version
        {
            get
            {
                if (ApplicationCommands.visio_app_version == null)
                {
                    var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);

                    var application = cmdtarget.Application;
                    ApplicationCommands.visio_app_version = VisioAutomation.Application.ApplicationHelper.GetVersion(application);
                }
                return ApplicationCommands.visio_app_version;
            }            
        }

        public VA.Application.UndoScope NewUndoScope(string name)
        {
            if (this.VisioApplication == null)
            {
                throw new System.ArgumentException("Cant create UndoScope. There is no visio application attached.");
            }

            return new VA.Application.UndoScope(this.VisioApplication, name);
        }
    }
}