using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
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
                var cmdtarget = this._client.GetCommandTargetApplication();

                var app = cmdtarget.Application;

                // if there's no active document, then there can't be an active document
                if (app.ActiveDocument == null)
                {
                    this._client.Output.WriteVerbose("HasActiveDocument: No Active Window");
                    return false;
                }

                var active_window = app.ActiveWindow;

                // If there's no active window there can't be an active document
                if (active_window == null)
                {
                    this._client.Output.WriteVerbose("HasActiveDocument: No Active Document");
                    return false;
                }

                // Check if the window type matches that of a document
                short active_window_type = active_window.Type;
                var vis_drawing = (int)IVisio.VisWinTypes.visDrawing;
                var vis_master = (int)IVisio.VisWinTypes.visMasterWin;
                // var vis_sheet = (short)IVisio.VisWinTypes.visSheet;

                this._client.Output.WriteVerbose("The Active Window: Type={0} & SybType={1}", active_window_type, active_window.SubType);
                if (!(active_window_type == vis_drawing || active_window_type == vis_master))
                {
                    this._client.Output.WriteVerbose("The Active Window Type must be one of {0} or {1}", IVisio.VisWinTypes.visDrawing, IVisio.VisWinTypes.visMasterWin);
                    return false;
                }

                //  verify there is an active page
                if (app.ActivePage == null)
                {
                    this._client.Output.WriteVerbose("HasActiveDocument: Active Page is null");

                    if (active_window.SubType == 64)
                    {
                        // 64 means master is being edited

                    }
                    else
                    {
                        this._client.Output.WriteVerbose("HasActiveDocument: Active Page is null");
                        return false;
                    }
                }

                this._client.Output.WriteVerbose("HasActiveDocument: Verified a drawing is available for use");

                return true;
            }
        }

        public void ActivateDocumentWithName(string name)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            var documents = cmdtarget.Application.Documents;
            var doc = documents[name];

            this.ActivateDocument(doc);
        }

        public void ActivateDocument(IVisio.Document doc)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            // if the doc is already active do nothing
            if (doc == cmdtarget.ActiveDocument)
            {
                // do nothing
                return;
            }

            // go through each window and check if it is assigned
            // to the target document
            var allwindows = cmdtarget.Application.Windows.ToEnumerable();
            var target_win = allwindows.FirstOrDefault(w => w.Document == doc);

            if (target_win == null)
            {
                // no window found
                throw new VisioOperationException("Could not find window for document");
            }

            target_win.Activate();
            if (cmdtarget.Application.ActiveDocument != doc)
            {
                // tried to activate window, but active document does not reflect it
                throw new InternalAssertionException("Failed to activate document");
            }
        }

        public void CloseDocument(TargetDocument targetdoc, bool force)
        {
            targetdoc = targetdoc.Resolve(this._client);


            var doc = targetdoc.Item;
            var app = doc.Application;

            if (doc.Type != IVisio.VisDocumentTypes.visTypeDrawing)
            {
                this._client.Output.WriteVerbose("Not a Drawing Window", doc.Name);
                throw new System.ArgumentException("Not a Drawing Window");
            }

            this._client.Output.WriteVerbose("Closing Document Name=\"{0}\"", doc.Name);
            this._client.Output.WriteVerbose("Closing Document FullName=\"{0}\"", doc.FullName);

            if (force)
            {
                using (var alert = new VisioAutomation.Application.AlertResponseScope(app, VisioAutomation.Application.AlertResponseCode.No))
                {
                    doc.Close();
                }
            }
            else
            {
                doc.Close();
            }
        }

        public void CloseDocuments(VisioScripting.TargetDocuments targetdocs, bool force)
        {
            this._client.Output.WriteVerbose("Closing {0} documents", targetdocs.Items.Count);
            foreach (var target_doc in targetdocs.Items)
            {
                this._client.Output.WriteVerbose("Closing doc with ID={0} Name={1}", target_doc.ID, target_doc.Name);
                target_doc.Close(force);
            }
        }

        public void CloseAllDocumentsWithoutSaving()
        {
            var cmdtarget = this._client.GetCommandTargetApplication();
            var documents = cmdtarget.Application.Documents;
            var docs = documents.ToEnumerable().Where(doc => doc.Type == IVisio.VisDocumentTypes.visTypeDrawing).ToList();

            using (var alert = new VisioAutomation.Application.AlertResponseScope(cmdtarget.Application, VisioAutomation.Application.AlertResponseCode.No))
            {
                foreach (var doc in docs)
                {
                    this._client.Output.WriteVerbose( "Closing Document Name=\"{0}\"", doc.Name);
                    this._client.Output.WriteVerbose( "Closing Document FullName=\"{0}\"", doc.FullName);
                    doc.Close();
                }
            }
        }

        public IVisio.Document NewDocument()
        {
            return this.NewDocumentFromTemplate(null);
        }

        public IVisio.Document NewDocumentFromTemplate(string template)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            this._client.Output.WriteVerbose("Creating Empty Drawing");
            var documents = cmdtarget.Application.Documents;
            
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

        public void SaveDocument(TargetDocument targetdoc)
        {
            targetdoc = targetdoc.Resolve(this._client);
            targetdoc.Item.Save();
        }


        public void SaveDocumentAs(TargetDocument targetdoc, string filename)
        {
            targetdoc = targetdoc.Resolve(this._client);
            targetdoc.Item.SaveAs(filename);
        }

        public IVisio.Document NewDocument(VisioAutomation.Geometry.Size size)
        {
            return this.NewDocumentFromTemplate(size,null);
        }

        public IVisio.Document NewDocumentFromTemplate(VisioAutomation.Geometry.Size size, string template)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            var doc = this.NewDocumentFromTemplate(template);
            var pagecells = new VisioAutomation.Pages.PageFormatCells();
            pagecells.Width = size.Width;
            pagecells.Height = size.Height;

            var pages = doc.Pages;
            var page = pages[1];

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(pagecells);
            writer.Commit(page.PageSheet, CellValueType.Formula);

            return doc;
        }

        public IVisio.Document OpenStencilDocument(string name)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length == 0)
            {
                throw new System.ArgumentException(nameof(name));
            }

            this._client.Output.WriteVerbose( "Loading stencil \"{0}\"", name);

            var documents = cmdtarget.Application.Documents;
            var doc = documents.OpenStencil(name);

            this._client.Output.WriteVerbose( "Finished loading stencil \"{0}\"", name);
            return doc;
        }

        public IVisio.Document OpenDocument(string filename)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }

            if (filename.Length == 0)
            {
                throw new System.ArgumentException("filename cannot be empty", nameof(filename));
            }

            string abs_filename = System.IO.Path.GetFullPath(filename);

            this._client.Output.WriteVerbose( "Input filename: {0}", filename);
            this._client.Output.WriteVerbose( "Absolute filename: {0}", abs_filename);

            if (!System.IO.File.Exists(abs_filename))
            {
                string msg = string.Format("File \"{0}\"does not exist", abs_filename);
                throw new System.ArgumentException(msg, nameof(filename));
            }

            var documents = cmdtarget.Application.Documents;
            var doc = documents.Add(filename);
            return doc;
        }

        public IVisio.Document GetDocumentWithName(string name)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            var documents = cmdtarget.Application.Documents;
            var doc = documents[name];
            return doc;
        }

        public List<IVisio.Document> FindDocuments(string namepattern, IVisio.VisDocumentTypes? doctype)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            var docs = cmdtarget.Application.Documents;

            // first get the full list
            var doc_list = docs.ToEnumerable().ToList();

            if (doctype == null)
            {
                return doc_list;
            }
            // then filter by doc types
            doc_list = doc_list.Where(d => doctype.Value == d.Type).ToList();

            // second perform any name filtering

            if (namepattern == null)
            {
                return doc_list;
            }

            var filter_action = VisioScripting.Helpers.WildcardHelper.FilterAction.Include;
            doc_list = VisioScripting.Helpers.WildcardHelper.FilterObjectsByNames(
                doc_list, 
                new[] {namepattern}, 
                d => d.Name, 
                true, 
                filter_action).ToList();
            return doc_list;
        }
    }
}