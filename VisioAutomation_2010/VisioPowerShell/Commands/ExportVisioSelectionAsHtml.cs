using System;
using System.IO;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Export, VisioPowerShell.Commands.Nouns.VisioSelectionAsHtml)]
    public class ExportVisioSelectionAsHtml : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        [SMA.ValidateNotNullOrEmpty]
        public string Filename;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Overwrite;

        protected override void ProcessRecord()
        {
            if (!File.Exists(this.Filename))
            {
                this.WriteVerbose("File already exists");
                if (this.Overwrite)
                {
                    File.Delete(this.Filename);
                }
                else
                {
                    string msg = string.Format("File \"{0}\" already exists", this.Filename);
                    var exc = new ArgumentException(msg);
                    throw exc;
                }
            }

            string ext = Path.GetExtension(this.Filename).ToLowerInvariant();

            if (ext == ".html" || ext == ".xhtml" || ext == ".htm")
            {
                this.Client.Export.SelectionToHtml(this.Filename);                
            }
            else
            {
                this.Client.Export.SelectionToFile(this.Filename);
            }
        }
    }
}