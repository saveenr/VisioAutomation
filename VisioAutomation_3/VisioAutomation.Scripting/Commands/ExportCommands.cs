using System.Linq;
using SXL=System.Xml.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ExportCommands : SessionCommands
    {
        public ExportCommands(Session session) :
            base(session)
        {

        }

        public void ExportPageToFile(string filename)
        {
            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var old_selection = this.Session.Selection.EnumSelectedShapes().ToList();

            this.Session.Selection.SelectNone();
            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;
            active_page.Export(filename);
            var active_window = application.ActiveWindow;
            active_window.Select(old_selection, IVisio.VisSelectArgs.visSelect);
        }

        public void ExportSelectionToFile(string filename)
        {
            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Session.Selection.GetSelection();
            selection.Export(filename);
        }

        public void ExportPagesToFiles(string filename)
        {
            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            var application = this.Session.VisioApplication;
            var old_page = application.ActivePage;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages.AsEnumerable().ToList();
            var pbase = System.IO.Path.GetDirectoryName(filename);

            if (!System.IO.Directory.Exists(pbase))
            {
                this.Session.Write(OutputStream.Error, " Folder {0} does not exist", pbase);
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

                    this.Session.Write(OutputStream.User, "file {0}", page_filname);
                    page_filname = System.IO.Path.Combine(pbase, page_filname);
                    page.Activate();
                    this.Session.Selection.SelectNone();
                    page.Export(page_filname);
                }
            old_page.Activate();
        }

        public void ExportSelectionAsSVGXHTML(string filename)
        {
            // Save temp SVG
            string svg_filename = SaveSelectionAsTemporarySVG();

            // Load temp SVG
            var load_svg_timer = new System.Diagnostics.Stopwatch();
            var svg_doc = SXL.XDocument.Load(svg_filename);
            load_svg_timer.Stop();
            this.Session.Write(OutputStream.Verbose, "Finished SVG Loading ({0} seconds)", load_svg_timer.Elapsed.TotalSeconds);

            // Delete temp SVG
            DeleteTemporarySVG(svg_filename);

            this.Session.Write(OutputStream.Verbose, "Creating XHTML with embedded SVG");
            var s = svg_filename;

            if (System.IO.File.Exists(filename))
            {
                this.Session.Write(OutputStream.Verbose, "Deleting \"{0}\"", filename);
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
            this.Session.Write(OutputStream.Verbose, "Done writing XHTML file \"{0}\"", filename);
        }


        public void DeleteTemporarySVG(string svg_filename)
        {
            if (System.IO.File.Exists(svg_filename))
            {
                this.Session.Write(OutputStream.Verbose, "Deleting \"{0}\"", svg_filename);
                System.IO.File.Delete(svg_filename);
            }
            else
            {
                string msg = string.Format("Expected to find temp svg filename \"{0}\" but it does not exist", svg_filename);
                this.Session.Write(OutputStream.Verbose, msg);
                throw new AutomationException(msg);
            }
        }

        public string SaveSelectionAsTemporarySVG()
        {
            string svg_filename = System.IO.Path.GetTempFileName() + "_temp.svg";

            if (System.IO.File.Exists(svg_filename))
            {
                System.IO.File.Delete(svg_filename);
            }

            var export_timer = new new System.Diagnostics.Stopwatch();

            this.Session.Write(OutputStream.Verbose, "Started SVG export to \"{0}\"", svg_filename);

            var selection = this.Session.Selection.GetSelection();
            selection.Export(svg_filename);
            export_timer.Stop();

            this.Session.Write(OutputStream.Verbose, "Finished SVG export ({0} seconds)", export_timer.Elapsed.TotalSeconds);

            var fi = new System.IO.FileInfo(svg_filename);
            this.Session.Write(OutputStream.Verbose, "SVG File size = {0} bytes", fi.Length);
            return svg_filename;
        }
    }
}

