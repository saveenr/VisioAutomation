using System.Management.Automation;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, "VisioDocument")]
    public class New_VisioDocument : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            if (!this.client.Application.HasApplication)
            {
                this.client.Application.New();
            }
            else
            {
                if (!this.client.Application.Validate())
                {
                    this.client.Application.New();
                }
            }
            var doc = this.client.Document.New();
            this.WriteObject(doc);
        }
    }
}