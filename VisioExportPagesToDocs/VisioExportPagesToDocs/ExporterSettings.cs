using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioExportPagesToDocs
{
    public class ExporterSettings
    {
        public IVisio.Document InputDocument;
        public string DestinationPath { get; set; }
        public string BaseName { get; set; }
        public string InputExtension { get; set; }
        public bool Overwrite { get; set; }

        public ExporterSettings()
        {
            this.Overwrite = false;
        }
    }


    public class LogRecord
    {
        public ExporterSettings Settings;
        public string PageName;
        public int PageIndex;
        public string OutputFilename;
        public bool OutputFileAlreadyExisted;
        public bool OutputFileWritten;
    }

}