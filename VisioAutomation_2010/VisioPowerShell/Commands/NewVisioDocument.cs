using SMA = System.Management.Automation;


namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioDocument)]
    public class NewVisioDocument : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            if (!this.Client.Application.HasActiveApplication)
            {
                this.Client.Application.NewActiveApplication();
            }
            else
            {
                if (!this.Client.Application.ValidateActiveApplication())
                {
                    this.Client.Application.NewActiveApplication();
                }
            }
            var doc = this.Client.Document.NewDocument();
            this.WriteObject(doc);
        }
    }
}