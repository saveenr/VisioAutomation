using System.Management.Automation;

namespace VisioPowerShell.Commands.Save
{
    [Cmdlet(VerbsData.Save, "VisioDocument")]
    public class Save_VisioDocument : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = false)]
        [ValidateNotNullOrEmpty]
        public string Filename;

        protected override void ProcessRecord()
        {
            if (this.Filename!=null)
            {
                this.client.Document.SaveAs(this.Filename);
            }
            else
            {
                this.client.Document.Save();
            }
        }
    }
}