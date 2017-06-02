using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioDocument)]
    public class NewVisioDocument : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            if (!this.Client.Application.HasApplication)
            {
                this.Client.Application.New();
            }
            else
            {
                if (!this.Client.Application.Validate())
                {
                    this.Client.Application.New();
                }
            }
            var doc = this.Client.Document.New();
            this.WriteObject(doc);
        }
    }
}