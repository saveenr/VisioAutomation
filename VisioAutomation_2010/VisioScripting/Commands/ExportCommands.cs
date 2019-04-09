using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using SXL = System.Xml.Linq;

namespace VisioScripting.Commands
{
    public class ExportCommands : CommandSet
    {
        internal ExportCommands(Client client)
            : base(client)
        {
        }

        public void ExportPageToImage(TargetActivePage targetpage, string filename)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

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

            var old_selected_shapes = this._client.Selection.GetShapes( new VisioScripting.TargetActiveSelection());

            var targetwindow = new VisioScripting.TargetWindow();
            this._client.Selection.SelectNone(targetwindow);
            var application = cmdtarget.Application;
            var active_page = application.ActivePage;
            active_page.Export(filename);
            var active_window = application.ActiveWindow;
            active_window.Select(old_selected_shapes, IVisio.VisSelectArgs.visSelect);
        }

        public void ExportAllPagesToImages(TargetActiveDocument targetdoc, string filename)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

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

                var targetwindow = new VisioScripting.TargetWindow();

                this._client.Output.WriteUser("file {0}", page_filname);
                page_filname = System.IO.Path.Combine(pbase, page_filname);
                active_window.Page = page;
                this._client.Selection.SelectNone(targetwindow);
                page.Export(page_filname);
            }
            active_window.Page = old_page;
        }

        public void ExportSelectionToImage(TargetActiveSelection targetselection, string filename)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

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

            selection.Export(filename);
        }

        public void ExportSelectionToHtml(TargetActiveSelection targetselection, string filename)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

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

            this._export_to_html(selection, filename, s => this._client.Output.WriteVerbose(s));
        }

        private void _export_to_html(IVisio.Selection selection, string filename, System.Action<string> export_log)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            // Save temp SVG
            string svg_filename = System.IO.Path.GetTempFileName() + "_temp.svg";
            selection.Export(svg_filename);

            // Load temp SVG
            var load_svg_timer = new System.Diagnostics.Stopwatch();
            var svg_doc = SXL.XDocument.Load(svg_filename);
            load_svg_timer.Stop();
            export_log(string.Format("Finished SVG Loading ({0} seconds)", load_svg_timer.Elapsed.TotalSeconds));

            // Delete temp SVG
            if (System.IO.File.Exists(svg_filename))
            {
                System.IO.File.Delete(svg_filename);
            }

            export_log("Creating XHTML with embedded SVG");

            if (System.IO.File.Exists(filename))
            {
                export_log(string.Format("Deleting \"{0}\"", filename));
                System.IO.File.Delete(filename);
            }

            var xhtml_doc = new SXL.XDocument();
            var xhtml_root = new SXL.XElement("{http://www.w3.org/1999/xhtml}html");
            xhtml_doc.Add(xhtml_root);
            var svg_node = svg_doc.Root;
            svg_node.Remove();

            var body = new SXL.XElement("{http://www.w3.org/1999/xhtml}body");
            xhtml_root.Add(body);
            body.Add(svg_node);

            xhtml_doc.Save(filename);
            export_log(string.Format("Done writing XHTML file \"{0}\"", filename));
        }
    }
}