using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsData.Export, VisioPowerShell.Nouns.VisioPage)]
    public class Export_VisioPage : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)] 
        [ValidateNotNullOrEmpty]
        public string Filename;

        [Parameter(Position = 1, Mandatory = false)]
        public SwitchParameter AllPages;

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