using System.Management.Automation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Redo
{
    [Cmdlet(VerbsCommon.Redo, "Visio")]
    public class Redo_Visio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.client.Application.Redo();
        }
    }
}