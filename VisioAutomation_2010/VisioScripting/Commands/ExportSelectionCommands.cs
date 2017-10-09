using System;
using IVisio = Microsoft.Office.Interop.Visio;
using SXL = System.Xml.Linq;

namespace VisioScripting.Commands
{
    public class ExportSelectionCommands : CommandSet
    {
        internal ExportSelectionCommands(Client client)
            : base(client)
        {
        }

        public void SelectionToFile(string filename)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            if (selection.Count < 1)
            {
                string msg = String.Format("Selection contains no shapes");
                throw new System.ArgumentException(msg);
            }

            selection.Export(filename);
        }

        public void SelectionToHtml(string filename)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            if (selection.Count<1)
            {
                string msg = String.Format("Selection contains no shapes");
                throw new System.ArgumentException(msg);
            }

            this.SelectionToHtml(selection, filename, s => this._client.Output.WriteVerbose(s));
        }

        private void SelectionToHtml(IVisio.Selection selection, string filename, System.Action<string> export_log)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

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