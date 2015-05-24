using System.Management.Automation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Undo
{
    [Cmdlet(SMA.VerbsCommon.Undo, "Visio")]
    public class Undo_Visio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.client.Application.Undo();
        }
    }
}