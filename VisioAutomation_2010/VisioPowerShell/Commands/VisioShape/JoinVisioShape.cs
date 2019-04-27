using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Join, Nouns.VisioShape)]
    public class JoinVisioShape : VisioCmdlet
    {
        // PSEUDOCONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            this.HandlePseudoContext(this.Shape);

            var group = this.Client.Grouping.Group(VisioScripting.TargetSelection.Auto);
            this.WriteObject(group);
        }
    }
}