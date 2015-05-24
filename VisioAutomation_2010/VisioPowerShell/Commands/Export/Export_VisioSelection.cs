using System;
using System.IO;
using System.Management.Automation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Export
{
    [Cmdlet(SMA.VerbsData.Export, "VisioSelectionAsXHTML")]
    public class Export_VisioSelectionAsXHTML : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string Filename;

        [Parameter(Mandatory = false)]
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
                this.client.Export.SelectionToSVGXHTML(this.Filename);                
            }
            else
            {
                this.client.Export.SelectionToFile(this.Filename);
            }
        }
    }
}