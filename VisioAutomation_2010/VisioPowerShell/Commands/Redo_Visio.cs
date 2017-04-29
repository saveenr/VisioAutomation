using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Redo, VisioPowerShell.Commands.Nouns.Visio)]
    public class Redo_Visio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Application.Redo();
        }
    }
}