using System.Linq;

using VisioAutomation.Extensions;

using IVisio = Microsoft.Office.Interop.Visio;
using SXL = System.Xml.Linq;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ExportCommands : CommandSet
    {
        public ExportCommands(Client client)
            : base(client)
        {
        }

        public void PageToFile(string filename)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!this.Client.Selection.HasShapes())
            {
                this.Client.WriteVerbose("No selected shapes. Not exporting.");
                return;
            }

            var old_selection = this.Client.Selection.GetShapes();

            this.Client.Selection.None();
            var application = this.Client.VisioApplication;
            var active_page = application.ActivePage;
            active_page.Export(filename);
            var active_window = application.ActiveWindow;
            active_window.Select(old_selection, IVisio.VisSelectArgs.visSelect);
        }

        public void SelectionToFile(string filename)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!this.Client.Selection.HasShapes())
            {
                this.Client.WriteVerbose("No selected shapes. Not exporting.");
                return;
            }

            var selection = this.Client.Selection.Get();
            selection.Export(filename);
        }

        public void PagesToFiles(string filename)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            var application = this.Client.VisioApplication;
            var old_page = application.ActivePage;
            var active_document = application.ActiveDocument;
            var active_window = application.ActiveWindow;

            var pages = active_document.Pages.AsEnumerable().ToList();
            var pbase = System.IO.Path.GetDirectoryName(filename);

            if (!System.IO.Directory.Exists(pbase))
            {
                this.Client.WriteError(" Folder {0} does not exist", pbase);
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
                string page_filname = System.String.Format(
                    "{0}_{1}_{2}{3}{4}",
                    fbase,
                    page_index,
                    page.Name,
                    bkgnd,
                    ext);

                this.Client.WriteUser("file {0}", page_filname);
                page_filname = System.IO.Path.Combine(pbase, page_filname);
                active_window.Page = page;
                this.Client.Selection.None();
                page.Export(page_filname);
            }
            active_window.Page = old_page;
        }

        public void SelectionToSVGXHTML(string filename)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!this.Client.Selection.HasShapes())
            {
                this.Client.WriteVerbose("No selected shapes. Not exporting.");
                return;
            }

            var selection = this.Client.Selection.Get();
            this.SelectionToSVGXHTML(this.Client.Selection.Get(), filename, s => this.Client.WriteVerbose(s));
        }

        private void SelectionToSVGXHTML(IVisio.Selection selection, string filename, System.Action<string> verboselog)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            // Save temp SVG
            string svg_filename = System.IO.Path.GetTempFileName() + "_temp.svg";
            selection.Export(svg_filename);

            // Load temp SVG
            var load_svg_timer = new System.Diagnostics.Stopwatch();
            var svg_doc = SXL.XDocument.Load(svg_filename);
            load_svg_timer.Stop();
            verboselog(string.Format("Finished SVG Loading ({0} seconds)", load_svg_timer.Elapsed.TotalSeconds));

            // Delete temp SVG
            if (System.IO.File.Exists(svg_filename))
            {
                System.IO.File.Delete(svg_filename);
            }
            else
            {
            }

            verboselog("Creating XHTML with embedded SVG");
            var s = svg_filename;

            if (System.IO.File.Exists(filename))
            {
                verboselog(string.Format("Deleting \"{0}\"", filename));
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
            verboselog(string.Format("Done writing XHTML file \"{0}\"", filename));
        }

        public void ExportSelectionToXAML(string filename)
        {
            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!this.Client.Selection.HasShapes())
            {
                return;
            }

            var selection = this.Client.Selection.Get();
            ExportSelectionAsXAML2(this.Client.Selection.Get(), filename, s => this.Client.Output.WriteVerbose(s));
        }

        public static void ExportSelectionAsXAML2(
            IVisio.Selection sel,
            string filename,
            System.Action<string> verboselog)
        {
            // Save temp SVG

            string svg_filename = System.IO.Path.GetTempFileName() + "_temp.svg";

            sel.Export(svg_filename);

            // Load temp SVG

            var load_svg_timer = new System.Diagnostics.Stopwatch();

            string input_svg_content = System.IO.File.ReadAllText(svg_filename);

            load_svg_timer.Stop();

            verboselog(string.Format("Finished SVG Loading ({0} seconds)", load_svg_timer.Elapsed.TotalSeconds));

            // Delete temp SVG

            if (System.IO.File.Exists(svg_filename))
            {
                System.IO.File.Delete(svg_filename);
            }
            else
            {
                string msg = string.Format("Temporary SVG file could not be found: \"{0}\"", svg_filename);
                throw new VisioOperationException(msg);
            }

            verboselog("Creating XHTML with embedded SVG");

            var s = svg_filename;

            if (System.IO.File.Exists(filename))
            {
                verboselog(string.Format("Deleting \"{0}\"", filename));

                System.IO.File.Delete(filename);
            }

            verboselog("Converting to XAML ...");

            var convert_timer = new System.Diagnostics.Stopwatch();

            string xaml;

            try
            {
                xaml = XamlTuneConverter.Svg2Xaml.ConvertFromSVG(input_svg_content);
            }

            catch (System.Exception e)
            {
                string msg = System.String.Format("Failed to convert to XAML \"{0}\"", e.Message + e.StackTrace);

                verboselog(msg);

                return;
            }

            convert_timer.Stop();

            verboselog("Writing XAML File");

            System.IO.File.WriteAllText(filename, xaml);

            verboselog("Finished writing XAML File");

            verboselog(string.Format("Finished XAML export ({0} seconds)", convert_timer.Elapsed.TotalSeconds));
        }
    }
}