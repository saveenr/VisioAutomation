using System.Linq;
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

        public void ExportPageToImage(TargetPage targetpage, string filename)
        {
            targetpage = targetpage.Resolve(this._client);

            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }

            targetpage.Page.Export(filename);
        }

        public void  ExportPagesToImages(TargetDocument targetdoc, string filename)
        {
            var output_folder = System.IO.Path.GetDirectoryName(filename);

            if (!System.IO.Directory.Exists(output_folder))
            {
                this._client.Output.WriteError(" Folder {0} does not exist", output_folder);
                return;
            }

            targetdoc = targetdoc.Resolve(this._client);

            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }

            var pages = targetdoc.Document.Pages.ToList();

            var ext = System.IO.Path.GetExtension(filename);
            string filename_base = System.IO.Path.GetFileNameWithoutExtension(filename);

            foreach (int page_index in Enumerable.Range(0,pages.Count))
            {
                var page = pages[page_index];
                string bkgnd_tag = "";
                if (page.Background != 0)
                {
                    bkgnd_tag = "(Background)";
                }
                string page_filname = string.Format("{0}_{1}_{2}_{3}{4}", filename_base, page_index, page.Name, bkgnd_tag, ext);

                page_filname = System.IO.Path.Combine(output_folder, page_filname);
                page.Export(page_filname);
            }
        }

        public void ExportSelectionToImage(TargetActiveSelection targetselection, string filename)
        {
            targetselection = targetselection.Resolve(this._client);

            if (filename == null)
            {
                throw new System.ArgumentNullException(nameof(filename));
            }

            targetselection.Selection.Export(filename);
        }

        public void ExportSelectionToHtml(TargetActiveSelection targetselection, string filename)
        {
            targetselection = targetselection.Resolve(this._client);

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