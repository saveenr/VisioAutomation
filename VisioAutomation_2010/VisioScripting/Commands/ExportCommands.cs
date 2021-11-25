
namespace VisioScripting.Commands
{
    public class ExportCommands : CommandSet
    {
        internal ExportCommands(Client client)
            : base(client)
        {
        }

        public void ExportPageToImage(TargetPage targetpage, string filename)
        {
            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }

            targetpage = targetpage.ResolveToPage(this._client);
            
            targetpage.Page.Export(filename);
        }


        public void ExportSelectionToImage(TargetSelection targetselection, string filename)
        {
            targetselection = targetselection.ResolveToSelection(this._client);

            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }

            targetselection.Selection.Export(filename);
        }

        public void ExportSelectionToHtml(TargetSelection targetselection, string filename)
        {
            targetselection = targetselection.ResolveToSelection(this._client);

            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }


            this._export_to_html(targetselection.Selection, filename, s => this._client.Output.WriteVerbose(s));
        }

        private void _export_to_html(IVisio.Selection selection, string filename, System.Action<string> export_log)
        {
     
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