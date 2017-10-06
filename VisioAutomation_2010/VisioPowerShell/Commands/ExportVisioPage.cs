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
            if (!this.AllPages)
            {
                // this means use the current page 
                this.Client.Export.PageToFile(this.Filename);
            }
            else
            {
                // is -AllPages is set then export them all
                this.Client.Export.PagesToFiles(this.Filename);
            }
        }
    }
}