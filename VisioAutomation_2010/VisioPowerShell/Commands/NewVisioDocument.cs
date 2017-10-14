using SMA = System.Management.Automation;


namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioDocument)]
    public class NewVisioDocument : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            if (!this.Client.Application.HasApplication)
            {
                this.Client.Application.NewApplication();
            }
            else
            {
                if (!this.Client.Application.ValidateApplication())
                {
                    this.Client.Application.NewApplication();
                }
            }
            var doc = this.Client.Document.NewDocument();
            this.WriteObject(doc);
        }
    }
}