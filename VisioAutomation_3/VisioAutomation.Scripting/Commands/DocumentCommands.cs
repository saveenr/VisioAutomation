using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VAS = VisioAutomation.Scripting;

namespace VisioAutomation.Scripting.Commands
{
    public class DocumentCommands : SessionCommands
    {
        public DocumentCommands(Session session) :
            base(session)
        {

        }

        public void CloseDocument(bool force)
        {
            if (!this.Session.HasActiveDrawing())
            {
                return;
            }

            var application = this.Session.Application;
            var doc = application.ActiveDocument;

            if (doc.Type == IVisio.VisDocumentTypes.visTypeDrawing)
            {
                this.Session.Write(OutputStream.Verbose, "Closing Document Name=\"{0}\"", doc.Name);
                this.Session.Write(OutputStream.Verbose, "Closing Document FullName=\"{0}\"", doc.FullName);

                if (force)
                {
                    using (var alert = application.CreateAlertResponseScope(VA.UI.AlertResponseCode.No))
                    {
                        doc.Close();
                    }
                }
                else
                {
                    doc.Close();
                }
            }
        }

        public void CloseAllDocumentsWithoutSaving()
        {
            var application = this.Session.Application;
            var documents = application.Documents;
            var docs = documents.AsEnumerable().Where(doc => doc.Type == IVisio.VisDocumentTypes.visTypeDrawing).
                ToList();

            using (var alert = application.CreateAlertResponseScope(VA.UI.AlertResponseCode.No))
            {
                foreach (var doc in docs)
                {
                    this.Session.Write(OutputStream.Verbose, "Closing Document Name=\"{0}\"", doc.Name);
                    this.Session.Write(OutputStream.Verbose, "Closing Document FullName=\"{0}\"", doc.FullName);
                    doc.Close();
                }
            }
        }

        public IVisio.Document NewDocument()
        {
            this.Session.Write(OutputStream.Verbose, "Creating Empty Drawing");
            var application = this.Session.Application;
            var documents = application.Documents;
            var doc = documents.Add(string.Empty);
            return doc;
        }

        public void SaveDocument()
        {
            if (!this.Session.HasActiveDrawing())
            {
                this.Session.Write(OutputStream.Error, "No Drawing to Save");
                return;
            }
            var application = this.Session.Application;
            var doc = application.ActiveDocument;
            doc.Save();
        }

        public void SaveDocumentAs(string filename)
        {
            if (!this.Session.HasActiveDrawing())
            {
                this.Session.Write(OutputStream.Error, "No Drawing to Save");
                return;
            }

            var application = this.Session.Application;
            var doc = application.ActiveDocument;
            doc.SaveAs(filename);
        }

        public IVisio.Document NewDocument(double w, double h)
        {
            var doc = NewDocument();
            var page = this.Session.Application.ActivePage;
            page.SetSize(w, h);
            return doc;
        }

        public IVisio.Document OpenStencil(string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(name);
            }

            if (name.Length == 0)
            {
                throw new System.ArgumentException(name);
            }

            this.Session.Write(OutputStream.Verbose, "Loading stencil \"{0}\"", name);

            var application = this.Session.Application;
            var documents = application.Documents;
            var doc = documents.OpenStencil(name);

            this.Session.Write(OutputStream.Verbose, "Finished loading stencil \"{0}\"", name);
            return doc;
        }

        public IVisio.Document NewStencil()
        {
            var application = this.Session.Application;
            var documents = application.Documents;
            var doc = VA.DocumentHelper.NewStencil(documents);
            return doc;
        }

        public IVisio.Document OpenDocument(string filename)
        {
            if (filename == null)
            {
                throw new System.ArgumentNullException(filename);
            }

            if (filename.Length == 0)
            {
                throw new System.ArgumentException(filename);
            }

            string abs_filename = System.IO.Path.GetFullPath(filename);

            this.Session.Write(OutputStream.Verbose, "Input filename: {0}", filename);
            this.Session.Write(OutputStream.Verbose, "Absolute filename: {0}", abs_filename);

            if (!System.IO.File.Exists(abs_filename))
            {
                throw new System.ArgumentException("File does not exist", "filename");
            }

            var application = this.Session.Application;
            var documents = application.Documents;
            var doc = documents.Add(filename);
            return doc;
        }


        public IVisio.Document GetDocument(string name)
        {
            var application = this.Session.Application;
            var documents = application.Documents;
            var doc = documents[name];
            return doc;
        }

    }
}