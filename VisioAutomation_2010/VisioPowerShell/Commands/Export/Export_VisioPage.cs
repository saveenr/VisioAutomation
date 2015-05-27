using System.Management.Automation;

namespace VisioPowerShell.Commands.Export
{
    [Cmdlet(VerbsData.Export, "VisioPage")]
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
                this.client.Export.PageToFile(this.Filename);
            }
            else
            {
                // is -AllPages is set then export them all
                this.client.Export.PagesToFiles(this.Filename);
            }
        }
    }
}