using System.Management.Automation;

namespace VisioPowerShell.Commands.Redo
{
    [Cmdlet(VerbsCommon.Redo, VisioPowerShell.Nouns.Visio)]
    public class Redo_Visio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Application.Redo();
        }
    }
}