using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsData.Export, Nouns.VisioShape)]
    public class ExportVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        [SMA.ValidateNotNullOrEmpty]
        public string Filename;


        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Overwrite;

        private static HashSet<string> _static_html_extensions;

        // CONTEXT:SHAPES - does not correctly use shape context
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            if (this.Filename == null)
            {
                throw new System.ArgumentNullException(nameof(this.Filename));
            }

            if (!System.IO.File.Exists(this.Filename))
            {
                this.WriteVerbose("File already exists");
                if (this.Overwrite)
                {
                    System.IO.File.Delete(this.Filename);
                }
                else
                {
                    string msg = string.Format("File \"{0}\" already exists", this.Filename);
                    var exc = new System.ArgumentException(msg);
                    throw exc;
                }
            }

            if (_static_html_extensions == null)
            {
                _static_html_extensions = new HashSet<string> { ".html", ".htm", ".xhtml" };
            }

            string ext = System.IO.Path.GetExtension(this.Filename).ToLowerInvariant();

            if (this.Shapes == null)
            {
                // use the active selection
            }
            else
            {
                if (this.Shapes.Length < 1)
                {
                    throw new System.ArgumentOutOfRangeException(nameof(this.Shapes), "Shapes parameter must contain at least one shape");
                }

                this.Client.Selection.SelectShapes(VisioScripting.TargetWindow.Auto, this.Shapes);
            }

            if (_static_html_extensions.Contains(ext))
            {
                this.Client.Export.ExportSelectionToHtml(VisioScripting.TargetSelection.Auto, this.Filename);                
            }
            else
            {
                this.Client.Export.ExportSelectionToImage(VisioScripting.TargetSelection.Auto, this.Filename);
            }
        }
    }
}