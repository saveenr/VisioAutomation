using System.Linq;
using SXL=System.Xml.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ExportCommands : CommandSet
    {
        public ExportCommands(Session session) :
            base(session)
        {

        }

        public void ExportPageToFile(string filename)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!this.Session.HasSelectedShapes())
            {
                this.Session.WriteVerbose("No selected shapes. Not exporting.");
                return;
            }

            var old_selection = this.Session.Selection.EnumShapes().ToList();

            this.Session.Selection.SelectNone();
            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;
            active_page.Export(filename);
            var active_window = application.ActiveWindow;
            active_window.Select(old_selection, IVisio.VisSelectArgs.visSelect);
        }

        public void ExportSelectionToFile(string filename)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!this.Session.HasSelectedShapes())
            {
                this.Session.WriteVerbose("No selected shapes. Not exporting.");
                return;
            }

            var selection = this.Session.Selection.Get();
            selection.Export(filename);
        }

        public void ExportPagesToFiles(string filename)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            var application = this.Session.VisioApplication;
            var old_page = application.ActivePage;
            var active_document = application.ActiveDocument;
            var active_window = application.ActiveWindow;

            var pages = active_document.Pages.AsEnumerable().ToList();
            var pbase = System.IO.Path.GetDirectoryName(filename);

            if (!System.IO.Directory.Exists(pbase))
            {
                this.Session.WriteError( " Folder {0} does not exist", pbase);
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
                    string page_filname = System.String.Format("{0}_{1}_{2}{3}{4}", fbase, page_index, page.Name,
                                                               bkgnd, ext);

                    this.Session.WriteUser( "file {0}", page_filname);
                    page_filname = System.IO.Path.Combine(pbase, page_filname);
                    active_window.Page = page;
                    this.Session.Selection.SelectNone();
                    page.Export(page_filname);
                }
            active_window.Page = old_page;
        }

        public void ExportSelectionToSVGXHTML(string filename)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!this.Session.HasSelectedShapes())
            {
                this.Session.WriteVerbose("No selected shapes. Not exporting.");
                return;
            }

            var selection = this.Session.Selection.Get();
            ExportSelectionToSVGXHTML(this.Session.Selection.Get(), filename, s => this.Session.WriteVerbose( s));
        }

        public void ExportSelectionToSVGXHTML(IVisio.Selection selection, string filename, System.Action<string> verboselog)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

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
                // TODO: throw an exception
            }

            verboselog(string.Format("Creating XHTML with embedded SVG"));
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
    }
}