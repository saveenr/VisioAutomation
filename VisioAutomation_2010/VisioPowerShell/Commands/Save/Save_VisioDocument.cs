using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsData.Save, "VisioDocument")]
    public class Save_VisioDocument : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = false)]
        [SMA.ValidateNotNullOrEmptyAttribute]
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