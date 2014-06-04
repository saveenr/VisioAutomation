using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Export, "VisioSelectionAsXHTML")]
    public class Export_VisioSelectionAsXHTML : VisioPS.VisioCmdlet
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
                this.WriteVerboseEx("File already exists");
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

            var scriptingsession = this.ScriptingSession;

            string ext = System.IO.Path.GetExtension(this.Filename).ToLowerInvariant();

            if (ext == ".html" || ext == ".xhtml" || ext == ".htm")
            {
                scriptingsession.Export.SelectionToSVGXHTML(this.Filename);                
            }
            else
            {
                scriptingsession.Export.SelectionToFile(this.Filename);
               
            }
        }
    }
}