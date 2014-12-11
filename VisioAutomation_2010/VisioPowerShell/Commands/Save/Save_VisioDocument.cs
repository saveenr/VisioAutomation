using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Save, "VisioDocument")]
    public class Save_VisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        [SMA.ValidateNotNullOrEmpty]
        public string Filename;

        protected override void ProcessRecord()
        {
            if (Filename!=null)
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