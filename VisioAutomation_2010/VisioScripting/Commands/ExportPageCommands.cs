using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class ExportPageCommands : CommandSet
    {
        internal ExportPageCommands(Client client)
            : base(client)
        {
        }

        public void ExportActicePageToFile(string filename)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            if (selection.Count < 1)
            {
                string msg = string.Format("Selection contains no shapes");
                throw new System.ArgumentException(msg);
            }

            var old_selected_shapes = this._client.Selection.GetShapesInSelection();

            this._client.Selection.SelectNone();
            var application = cmdtarget.Application;
            var active_page = application.ActivePage;
            active_page.Export(filename);
            var active_window = application.ActiveWindow;
            active_window.Select(old_selected_shapes, IVisio.VisSelectArgs.visSelect);
        }

        public void ExportActiveDocumentPagesToFiles(string filename)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }

            var application = cmdtarget.Application;
            var old_page = application.ActivePage;
            var active_document = application.ActiveDocument;
            var active_window = application.ActiveWindow;

            var pages = active_document.Pages.ToList();
            var pbase = System.IO.Path.GetDirectoryName(filename);

            if (!System.IO.Directory.Exists(pbase))
            {
                this._client.Output.WriteError(" Folder {0} does not exist", pbase);
                return;
            }
            var ext = System.IO.Path.GetExtension(filename);
            string fbase = System.IO.Path.GetFileNameWithoutExtension(filename);

            for (int page_index = 0; page_index < pages.Count; page_index++)
            {
                var page = pages[page_index];
                string bkgnd = "";
                if (page.Background != 0)
                {
                    bkgnd = "(Background)";
                }
                string page_filname = string.Format("{0}_{1}_{2}{3}{4}", fbase, page_index, page.Name, bkgnd, ext);

                this._client.Output.WriteUser("file {0}", page_filname);
                page_filname = System.IO.Path.Combine(pbase, page_filname);
                active_window.Page = page;
                this._client.Selection.SelectNone();
                page.Export(page_filname);
            }
            active_window.Page = old_page;
        }
    }
}