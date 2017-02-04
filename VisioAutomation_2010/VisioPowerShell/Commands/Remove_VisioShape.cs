using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Remove, VisioPowerShell.Nouns.VisioShape)]
    public class Remove_VisioShape : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            if (this.Shapes == null)
            {
                this.Client.Selection.Delete();                
            }
            else
            {
                foreach (var shape in this.Shapes)
                {
                    shape.Delete();
                }
            }
        }
    }
}