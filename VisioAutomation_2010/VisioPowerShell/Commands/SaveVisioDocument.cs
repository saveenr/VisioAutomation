using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsData.Save, VisioPowerShell.Commands.Nouns.VisioDocument)]
    public class SaveVisioDocument : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = false)]
        [ValidateNotNullOrEmpty]
        public string Filename;

        protected override void ProcessRecord()
        {
            if (this.Filename!=null)
            {
                this.Client.Document.SaveAs(this.Filename);
            }
            else
            {
                this.Client.Document.Save();
            }
        }
    }
}