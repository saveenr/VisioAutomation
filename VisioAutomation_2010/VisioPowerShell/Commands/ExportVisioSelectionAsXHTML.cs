using System;
using System.IO;
using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsData.Export, VisioPowerShell.Commands.Nouns.VisioSelectionAsXHTML)]
    public class ExportVisioSelectionAsXhtml : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string Filename;

        [Parameter(Mandatory = false)]
        public SwitchParameter Overwrite;

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