using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.Extensions;
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
                var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);

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

        public void Activate(string name)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);

            var documents = cmdtarget.Application.Documents;
            var doc = documents[name];

            this.Activate(doc);
        }

        public void Activate(IVisio.Document doc)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);


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

        public void Close(bool force)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            var doc = cmdtarget.ActiveDocument;

            if (doc.Type != IVisio.VisDocumentTypes.visTypeDrawing)
            {
                this._client.Output.WriteVerbose("Not a Drawing Window", doc.Name);
                throw new System.ArgumentException("Not a Drawing Window");
            }

            this._client.Output.WriteVerbose( "Closing Document Name=\"{0}\"", doc.Name);
            this._client.Output.WriteVerbose( "Closing Document FullName=\"{0}\"", doc.FullName);

            if (force)
            {
                using (var alert = new VisioAutomation.Application.AlertResponseScope(cmdtarget.Application, VisioAutomation.Application.AlertResponseCode.No))
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
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);
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

        public IVisio.Document New()
        {
            return this.NewWithTemplate(null);
        }

        public IVisio.Document NewWithTemplate(string template)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application);

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

        public void Save()
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);
            cmdtarget.ActiveDocument.Save();
        }

        public void SaveAs(string filename)
        {
            var cmdtarget = this._client.GetCommandTarget(CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);
            cmdtarget.ActiveDocument.SaveAs(filename);
        }

        public IVisio.Document New(VisioAutomation.Geometry.Size size)
        {
            return this.NewWithTemplate(size,null);
        }

        public IVisio.Document NewWithTemplate(VisioAutomation.Geometry.Size size,string template)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application);
            var doc = this.NewWithTemplate(template);
            this._client.Page.SetSize(size);
            return doc;
        }

        public IVisio.Document OpenStencil(string name)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application);

            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length == 0)
            {
                throw new System.ArgumentException("name");
            }

            this._client.Output.WriteVerbose( "Loading stencil \"{0}\"", name);

            var documents = cmdtarget.Application.Documents;
            var doc = documents.OpenStencil(name);

            this._client.Output.WriteVerbose( "Finished loading stencil \"{0}\"", name);
            return doc;
        }

        public IVisio.Document Open(string filename)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application);

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


        public IVisio.Document Get(string name)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application);

            var documents = cmdtarget.Application.Documents;
            var doc = documents[name];
            return doc;
        }

        public List<IVisio.Document> GetDocumentsByName(string name)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application);

            var documents = cmdtarget.Application.Documents;
            if (name == null || name == "*")
            {
                // return all documents
                var docs1 = documents.ToEnumerable().ToList();
                return docs1;
            }

            // get the named document
            var filter_action = VisioScripting.Helpers.WildcardHelper.FilterAction.Include;
            var docs2 = VisioScripting.Helpers.WildcardHelper.FilterObjectsByNames(documents.ToEnumerable(), new[] {name}, d => d.Name, true, filter_action).ToList();
            return docs2;
        }
    }
}