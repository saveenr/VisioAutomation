using System;
using System.IO;
using System.Management.Automation;

namespace VisioPowerShell.Commands.Export
{
    [Cmdlet(VerbsData.Export, "VisioSelectionAsXHTML")]
    public class Export_VisioSelectionAsXHTML : VisioCmdlet
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
                    string msg = $"File \"{this.Filename}\" already exists";
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