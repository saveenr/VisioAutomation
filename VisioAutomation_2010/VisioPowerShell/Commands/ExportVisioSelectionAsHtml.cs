using System;
using System.Collections.Generic;
using System.IO;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Export, VisioPowerShell.Commands.Nouns.VisioSelection)]
    public class ExportVisioSelection : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        [SMA.ValidateNotNullOrEmpty]
        public string Filename;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Overwrite;

        private static HashSet<string> _htmlExtensions;

        protected override void ProcessRecord()
        {
            if (this.Filename == null)
            {
                throw new System.ArgumentNullException(nameof(this.Filename));
            }
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


            if (_htmlExtensions == null)
            {
                _htmlExtensions = new HashSet<string> { ".html", ".htm", ".xhtml" };
            }

            string ext = Path.GetExtension(this.Filename).ToLowerInvariant();

            if (_htmlExtensions.Contains(ext))
            {
                this.Client.ExportSelection.SelectionToHtml(this.Filename);                
            }
            else
            {
                this.Client.ExportSelection.SelectionToFile(this.Filename);
            }
        }
    }
}