using VisioScripting;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsData.Export, Nouns.VisioPage)]
    public class ExportVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)] 
        [SMA.ValidateNotNullOrEmpty]
        public string Filename;

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public SMA.SwitchParameter AllPages;

        protected override void ProcessRecord()
        {
            if (this.AllPages)
            {
                this.Client.Export.ExportAllPagesToImages(new TargetActiveDocument(), this.Filename);
            }
            else
            {
                this.Client.Export.ExportPageToImage(new TargetActivePage(), this.Filename);
            }
        }
    }
}