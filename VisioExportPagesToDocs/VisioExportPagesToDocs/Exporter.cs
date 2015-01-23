using System.Linq;
using VisioAutomation.Extensions;

namespace VisioExportPagesToDocs
{
    public class Exporter
    {
        public ExporterSettings Settings;

        public Exporter(ExporterSettings settings)
        {
            if (settings == null)
            {
                throw new System.ArgumentNullException();
            }

            if (settings.InputDocument == null)
            {
                throw new System.ArgumentNullException();                
            }

            if (settings.DestinationPath == null)
            {
                throw new System.ArgumentNullException();
            }

            if (!System.IO.Directory.Exists(this.Settings.DestinationPath))
            {
                throw new System.ArgumentException(string.Format("destination path does not exist: \"{0}\"", this.Settings.DestinationPath));
            }

            string input_filname = settings.InputDocument.Name;

            this.Settings = settings;
            
            this.Settings.BaseName = this.Settings.BaseName ?? System.IO.Path.GetFileNameWithoutExtension(input_filname);
            this.Settings.InputExtension = System.IO.Path.GetExtension(input_filname);
        }

        public void Run()
        {            
            var pages = this.Settings.InputDocument.Pages;
            var app = this.Settings.InputDocument.Application;
            var docs = app.Documents;

            var _pages = pages.AsEnumerable().ToList();

            int pageindex = 1;
            foreach (var page in _pages)
            {
                string pagename = page.Name;
                var newdoc = docs.Add("");
                var newpage = newdoc.Pages[1];
                VisioAutomation.Pages.PageHelper.DuplicateToDocument(page, newdoc, newpage, pagename, true);

                // Visio allows characters in page names that are not valid for file names. Replace them.   
                foreach (var c in System.IO.Path.GetInvalidFileNameChars())
                {
                    pagename = pagename.Replace(c, '_');
                }

                string destname = System.IO.Path.Combine( this.Settings.DestinationPath,
                    this.Settings.BaseName + "_" + pageindex.ToString() + "_" + pagename + this.Settings.InputExtension);

                if (System.IO.File.Exists(destname))
                {
                    System.Console.WriteLine("Output file already exists. Skipping. File = \"{0}\"", destname);
                }

                var activewindow = app.ActiveWindow;
                activewindow.ViewFit = (int)Microsoft.Office.Interop.Visio.VisWindowFit.visFitPage;

                newdoc.SaveAs(destname);
                newdoc.Close(true);
                pageindex++;
            }
            this.Settings.InputDocument.Close(true);
        }
    }
}