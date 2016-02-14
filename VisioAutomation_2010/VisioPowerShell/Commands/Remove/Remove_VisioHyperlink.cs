using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Remove
{
    [Cmdlet(VerbsCommon.Remove, VisioPowerShell.Nouns.VisioHyperlink)]
    public class Remove_VisioHyperlink : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public int HyperlinkIndex { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            this.Client.Hyperlink.Delete(this.Shapes,this.HyperlinkIndex);
        }
    }
}