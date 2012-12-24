using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class DocumentCommands : CommandSet
    {
        public DocumentCommands(Session session) :
            base(session)
        {

        }

        public void Close(bool force)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var application = this.Session.VisioApplication;
            var doc = application.ActiveDocument;

            if (doc.Type == IVisio.VisDocumentTypes.visTypeDrawing)
            {
                this.Session.Write(OutputStream.Verbose, "Closing Document Name=\"{0}\"", doc.Name);
                this.Session.Write(OutputStream.Verbose, "Closing Document FullName=\"{0}\"", doc.FullName);

                if (force)
                {
                    using (var alert = application.CreateAlertResponseScope(VA.Application.AlertResponseCode.No))
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

        public void CloseAllWithoutSaving()
        {
            var application = this.Session.VisioApplication;
            var documents = application.Documents;
            var docs = documents.AsEnumerable().Where(doc => doc.Type == IVisio.VisDocumentTypes.visTypeDrawing).
                ToList();

            using (var alert = application.CreateAlertResponseScope(VA.Application.AlertResponseCode.No))
            {
                foreach (var doc in docs)
                {
                    this.Session.Write(OutputStream.Verbose, "Closing Document Name=\"{0}\"", doc.Name);
                    this.Session.Write(OutputStream.Verbose, "Closing Document FullName=\"{0}\"", doc.FullName);
                    doc.Close();
                }
            }
        }

        public IVisio.Document New()
        {
            this.Session.Write(OutputStream.Verbose, "Creating Empty Drawing");
            var application = this.Session.VisioApplication;
            var documents = application.Documents;
            var doc = documents.Add(string.Empty);
            return doc;
        }

        public void Save()
        {
            if (!this.Session.HasActiveDrawing)
            {
                this.Session.Write(OutputStream.Error, "No Drawing to Save");
                return;
            }
            var application = this.Session.VisioApplication;
            var doc = application.ActiveDocument;
            doc.Save();
        }

        public void SaveAs(string filename)
        {
            if (!this.Session.HasActiveDrawing)
            {
                this.Session.Write(OutputStream.Error, "No Drawing to Save");
                return;
            }

            var application = this.Session.VisioApplication;
            var doc = application.ActiveDocument;
            doc.SaveAs(filename);
        }

        public IVisio.Document New(double w, double h)
        {
            var doc = New();
            var page = this.Session.VisioApplication.ActivePage;

            var pagesize = new VA.Drawing.Size(w, h);
            VA.Pages.PageHelper.SetSize(page,pagesize);
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

            var application = this.Session.VisioApplication;
            var documents = application.Documents;
            var doc = documents.OpenStencil(name);

            this.Session.Write(OutputStream.Verbose, "Finished loading stencil \"{0}\"", name);
            return doc;
        }

        public IVisio.Document NewStencil()
        {
            var application = this.Session.VisioApplication;
            var documents = application.Documents;
            var doc = documents.AddEx(string.Empty, IVisio.VisMeasurementSystem.visMSDefault,
                          (int)IVisio.VisOpenSaveArgs.visAddStencil +
                          (int)IVisio.VisOpenSaveArgs.visOpenDocked,
                          0);
            return doc;
        }

        public IVisio.Document Open(string filename)
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

            var application = this.Session.VisioApplication;
            var documents = application.Documents;
            var doc = documents.Add(filename);
            return doc;
        }


        public IVisio.Document Get(string name)
        {
            var application = this.Session.VisioApplication;
            var documents = application.Documents;
            var doc = documents[name];
            return doc;
        }

    }
}