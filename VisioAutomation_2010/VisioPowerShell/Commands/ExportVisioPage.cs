using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Export, VisioPowerShell.Commands.Nouns.VisioPage)]
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
                this.Client.ExportPage.ExportAllPagesInActiveDocumentToFiles(this.Filename);
            }
            else
            {
                this.Client.ExportPage.ExportActivePageToFile(this.Filename);
            }
        }
    }
}