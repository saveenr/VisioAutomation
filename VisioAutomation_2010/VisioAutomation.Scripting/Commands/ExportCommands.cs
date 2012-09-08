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
            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!this.Session.HasSelectedShapes())
            {
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
            if (filename == null)
            {
                throw new System.ArgumentNullException("filename");
            }

            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Session.Selection.Get();
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
            VA.ExportHelper.ExportSelectionAsSVGXHTML2(this.Session.Selection.Get(),filename, s=>this.Session.Write(OutputStream.Verbose,s));
        }
    }
}

