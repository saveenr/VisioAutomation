using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;

namespace VisioExportPagesToDocs
{
    public class Exporter
    {
        public ExporterSettings Settings;
        public List<LogRecord> Log;

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

        public IEnumerable< LogRecord> Run()
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

                string basename = string.Format("{0}_{1}_{2}{3}", this.Settings.BaseName, pageindex.ToString(), pagename, this.Settings.InputExtension);
                string destname = System.IO.Path.Combine( this.Settings.DestinationPath, basename);

                var rec = new LogRecord();
                rec.OutputFileAlreadyExisted = System.IO.File.Exists(destname);
                rec.OutputFilename = destname;
                rec.Settings = this.Settings;
                rec.PageIndex = pageindex;
                rec.PageName = pagename;

                bool perform_save = false;
                if (this.Settings.Overwrite)
                {
                    perform_save = true;
                    if (rec.OutputFileAlreadyExisted)
                    {
                        System.IO.File.Delete(destname);
                    }                    
                }
                else
                {
                    perform_save = !rec.OutputFileAlreadyExisted;
                }

                if (perform_save)
                {
                    var activewindow = app.ActiveWindow;
                    activewindow.ViewFit = (int)Microsoft.Office.Interop.Visio.VisWindowFit.visFitPage;
                    newdoc.SaveAs(destname);
                    newdoc.Close(true);
                    rec.OutputFileWritten = true;
                }

                this.Log.Add(rec);

                yield return rec;
                pageindex++;
            }
            this.Settings.InputDocument.Close(true);
        }
    }
}