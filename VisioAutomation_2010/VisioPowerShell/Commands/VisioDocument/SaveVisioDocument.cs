using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Save, VisioPowerShell.Commands.Nouns.VisioDocument)]
    public class SaveVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        [SMA.ValidateNotNullOrEmpty]
        public string Filename;

        protected override void ProcessRecord()
        {
            if (this.Filename!=null)
            {
                this.Client.Document.SaveActiveDocumentAs(this.Filename);
            }
            else
            {
                this.Client.Document.SaveActiveDocument();
            }
        }
    }
}