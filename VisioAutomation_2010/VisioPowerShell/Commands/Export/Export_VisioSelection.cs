using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Export, "VisioSelectionAsXHTML")]
    public class Export_VisioSelectionAsXHTML : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        [SMA.ValidateNotNullOrEmpty]
        public string Filename;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Overwrite;

        protected override void ProcessRecord()
        {
            if (!System.IO.File.Exists(this.Filename))
            {
                this.WriteVerbose("File already exists");
                if (this.Overwrite)
                {
                    System.IO.File.Delete(this.Filename);
                }
                else
                {
                    var exc = new System.ArgumentException("File already exists");
                    throw exc;
                }
            }

            string ext = System.IO.Path.GetExtension(this.Filename).ToLowerInvariant();

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