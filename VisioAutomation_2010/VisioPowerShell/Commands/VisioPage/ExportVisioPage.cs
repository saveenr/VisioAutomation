using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsData.Export, Nouns.VisioPage)]
    public class ExportVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)] 
        [SMA.ValidateNotNullOrEmpty]
        public string Filename;


        //TODO: Instead of this parameter just identify all the pages explicitly
        [SMA.Parameter(Position = 1, Mandatory = false)]
        public SMA.SwitchParameter AllPages;

        protected override void ProcessRecord()
        {
            if (this.AllPages)
            {
                this.Client.Export.ExportPagesToImages(VisioScripting.TargetDocument.Auto, this.Filename);
            }
            else
            {
                this.Client.Export.ExportPageToImage(VisioScripting.TargetPage.Auto, this.Filename);
            }
        }
    }
}