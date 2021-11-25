using VisioAutomation.Extensions;


namespace VisioScripting.Commands
{
    public class ApplicationCommands : CommandSet
    {
        private static System.Version _static_visio_app_version;

        private IVisio.Application _app;
        
        internal ApplicationCommands(Client client) :
            this(client, null)
        {
        }

        internal ApplicationCommands(Client client, IVisio.Application app) :
            base(client)
        {
            this._app = app;
        }

        public bool HasApplication
        {
            get
            {
                bool b = this._app != null;
                this._client.Output.WriteVerbose("HasApplication: {0}", b);
                return b;
            }
        }

        public IVisio.Application GetApplication()
        {
            return this._app;
        }

        public void AssertHasApplication()
        {
            var has_app = this._client.Application.HasApplication;
            if (!has_app)
            {
                throw new System.ArgumentException("No Visio Application available");
            }
        }

        public void CloseApplication()
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireApplication);

            var app = cmdtarget.Application;

            if (app == null)
            {
                this._client.Output.WriteWarning("There is no Visio Application to stop");
                return;
            }

            // Force the app to close
            // it will require closing all documents and then quiting
            var documents = app.Documents;

            while (documents.Count > 0)
            {
                var active_document = app.ActiveDocument;
                active_document.Close(true); // force the document to close
            }

            // now that all the documents are closed, close the application
            app.Quit(true);

            // clean out the app reference
            this._app = null;
        }

        public IVisio.Application NewApplication()
        {
            this._client.Output.WriteVerbose("Creating a new Instance of Visio");
            var app = new IVisio.Application();
            this._client.Output.WriteVerbose("Attaching that instance to current scripting client");
            this._app = app;
            return app;
        }

        public bool ValidateApplication()
        {
            if (this._app == null)
            {
                this._client.Output.WriteVerbose("Client's Application object is null");
                return false;
            }

            try
            {
                // try to do something simple, read-only, and fast with the application object
                //  if No COMException was thrown when reading ProductName property. This application instance is treated as valid

                var app_version = this._app.ProductName;
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
                if (ApplicationCommands._static_visio_app_version == null)
                {
                    var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireApplication);

                    var application = cmdtarget.Application;
                    ApplicationCommands._static_visio_app_version = VisioAutomation.Application.ApplicationHelper.GetVersion(application);
                }
                return ApplicationCommands._static_visio_app_version;
            }            
        }


        public void MoveWindowToFront()
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireApplication);


            var app = cmdtarget.Application;

            if (app == null)
            {
                return;
            }

            VisioAutomation.Application.ApplicationHelper.BringWindowToTop(app);
        }

        public System.Drawing.Rectangle GetWindowRectangle()
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireApplication);


            var appwindow = cmdtarget.Application.Window;
            var rect = appwindow.GetWindowRect();
            return rect;
        }

        public void SetWindowRectangle( System.Drawing.Rectangle rect)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.RequireApplication);


            var appwindow = cmdtarget.Application.Window;
            appwindow.SetWindowRect(rect);
        }

        public void DeleteShapes(VisioScripting.TargetShapes targetshapes)
        {
            if (targetshapes.Resolved)
            {
                foreach (var shape in targetshapes.Shapes)
                {
                    shape.Delete();
                }
            }
            else
            {
                this._client.Selection.DeleteShapes(VisioScripting.TargetSelection.Auto);
            }
        }
    }
}