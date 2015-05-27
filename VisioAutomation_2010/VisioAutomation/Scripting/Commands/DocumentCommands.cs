using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class DocumentCommands : CommandSet
    {
        internal DocumentCommands(Client client) :
            base(client)
        {

        }

        public bool HasActiveDocument
        {
            get
            {
                var app = this.Client.Application.Get();

                // if there's no active document, then there can't be an active document
                if (app.ActiveDocument == null)
                {
                    this.Client.WriteVerbose("HasActiveDocument: No Active Window");
                    return false;
                }

                var active_window = app.ActiveWindow;

                // If there's no active window there can't be an active document
                if (active_window == null)
                {
                    this.Client.WriteVerbose("HasActiveDocument: No Active Document");
                    return false;
                }

                // Check if the window type matches that of a document
                short active_window_type = active_window.Type;
                var vis_drawing = (int)IVisio.VisWinTypes.visDrawing;
                var vis_master = (int)IVisio.VisWinTypes.visMasterWin;
                // var vis_sheet = (short)IVisio.VisWinTypes.visSheet;

                this.Client.WriteVerbose("The Active Window: Type={0} & SybType={1}", active_window_type, active_window.SubType);
                if (!(active_window_type == vis_drawing || active_window_type == vis_master))
                {
                    this.Client.WriteVerbose("The Active Window Type must be one of {0} or {1}", IVisio.VisWinTypes.visDrawing, IVisio.VisWinTypes.visMasterWin);
                    return false;
                }

                //  verify there is an active page
                if (app.ActivePage == null)
                {
                    this.Client.WriteVerbose("HasActiveDocument: Active Page is null");

                    if (active_window.SubType == 64)
                    {
                        // 64 means master is being edited

                    }
                    else
                    {
                        this.Client.WriteVerbose("HasActiveDocument: Active Page is null");
                        return false;
                    }
                }

                this.Client.WriteVerbose("HasActiveDocument: Verified a drawing is available for use");

                return true;
            }
        }

        internal void AssertDocumentAvailable()
        {
            if (!this.Client.Document.HasActiveDocument)
            {
                throw new VisioOperationException("No Drawing available");
            }
        }


        public void Activate(string name)
        {
            this.Client.Application.AssertApplicationAvailable();

            var application = this.Client.Application.Get();
            var documents = application.Documents;
            var doc = documents[name];

            this.Activate(doc);
        }

        public void Activate(IVisio.Document doc)
        {
            this.Client.Application.AssertApplicationAvailable();
            Documents.DocumentHelper.Activate(doc);
        }

        public void Close(bool force)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            var doc = application.ActiveDocument;

            if (doc.Type != IVisio.VisDocumentTypes.visTypeDrawing)
            {
                this.Client.WriteVerbose("Not a Drawing Window", doc.Name);
                throw new AutomationException("Not a Drawing Window");
            }

            this.Client.WriteVerbose( "Closing Document Name=\"{0}\"", doc.Name);
            this.Client.WriteVerbose( "Closing Document FullName=\"{0}\"", doc.FullName);

            if (force)
            {
                using (var alert = new Application.AlertResponseScope(application, Application.AlertResponseCode.No))
                {
                    doc.Close();
                }
            }
            else
            {
                doc.Close();
            }
        }

        public void CloseAllWithoutSaving()
        {
            this.Client.Application.AssertApplicationAvailable();
            var application = this.Client.Application.Get();
            var documents = application.Documents;
            var docs = documents.AsEnumerable().Where(doc => doc.Type == IVisio.VisDocumentTypes.visTypeDrawing).ToList();

            using (var alert = new Application.AlertResponseScope(application, Application.AlertResponseCode.No))
            {
                foreach (var doc in docs)
                {
                    this.Client.WriteVerbose( "Closing Document Name=\"{0}\"", doc.Name);
                    this.Client.WriteVerbose( "Closing Document FullName=\"{0}\"", doc.FullName);
                    doc.Close();
                }
            }
        }

        public IVisio.Document New()
        {
            return this.New(null);
        }

        public IVisio.Document New(string template)
        {
            this.Client.Application.AssertApplicationAvailable();

            this.Client.WriteVerbose("Creating Empty Drawing");
            var application = this.Client.Application.Get();
            var documents = application.Documents;
            
            if (template == null)
            {
                var doc = documents.Add(string.Empty);
                return doc;
            }
            else
            {

                var doc = documents.Add(string.Empty);
                var template_doc = documents.AddEx(template, IVisio.VisMeasurementSystem.visMSDefault,
                              (int)IVisio.VisOpenSaveArgs.visAddStencil +
                              (int)IVisio.VisOpenSaveArgs.visOpenDocked,
                              0);
                return doc;
            }
        }

        public void Save()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            var doc = application.ActiveDocument;
            doc.Save();
        }

        public void SaveAs(string filename)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            var doc = application.ActiveDocument;
            doc.SaveAs(filename);
        }

        public IVisio.Document New(double w, double h)
        {
            return this.New(w, h, null);
        }

        public IVisio.Document New(double w, double h,string template)
        {
            this.Client.Application.AssertApplicationAvailable();

            var doc = this.New(template);
            var pagesize = new Drawing.Size(w, h);
            this.Client.Page.SetSize(pagesize);
            return doc;
        }

        public IVisio.Document OpenStencil(string name)
        {
            this.Client.Application.AssertApplicationAvailable();
            
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length == 0)
            {
                throw new System.ArgumentException("name");
            }

            this.Client.WriteVerbose( "Loading stencil \"{0}\"", name);

            var application = this.Client.Application.Get();
            var documents = application.Documents;
            var doc = documents.OpenStencil(name);

            this.Client.WriteVerbose( "Finished loading stencil \"{0}\"", name);
            return doc;
        }

        public IVisio.Document Open(string filename)
        {
            this.Client.Application.AssertApplicationAvailable();
            
            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }

            if (filename.Length == 0)
            {
                throw new System.ArgumentException("filename cannot be empty", nameof(filename));
            }

            string abs_filename = System.IO.Path.GetFullPath(filename);

            this.Client.WriteVerbose( "Input filename: {0}", filename);
            this.Client.WriteVerbose( "Absolute filename: {0}", abs_filename);

            if (!System.IO.File.Exists(abs_filename))
            {
                string msg = $"File \"{abs_filename}\"does not exist";
                throw new System.ArgumentException(msg, nameof(filename));
            }

            var application = this.Client.Application.Get();
            var documents = application.Documents;
            var doc = documents.Add(filename);
            return doc;
        }


        public IVisio.Document Get(string name)
        {
            this.Client.Application.AssertApplicationAvailable();

            var application = this.Client.Application.Get();
            var documents = application.Documents;
            var doc = documents[name];
            return doc;
        }

        public List<IVisio.Document> GetDocumentsByName(string name)
        {
            var application = this.Client.Application.Get();
            var documents = application.Documents;
            if (name == null || name == "*")
            {
                // return all documents
                var docs1 = documents.AsEnumerable().ToList();
                return docs1;
            }

            // get the named document
            var docs2 = TextUtil.FilterObjectsByNames(documents.AsEnumerable(), new[] {name}, d => d.Name, true, TextUtil.FilterAction.Include).ToList();
            return docs2;
        }
    }
}