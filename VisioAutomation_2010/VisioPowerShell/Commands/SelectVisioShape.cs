using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Select, VisioPowerShell.Commands.Nouns.VisioShape)]
    public class SelectVisioShape : VisioCmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapes")]
        public IVisio.Shape[] Shapes { get; set; }
        
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapeIDs")]
        public int[] ShapeIDs { get; set; }
        
        [Parameter(Mandatory = true, Position=0, ParameterSetName = "SelectByOperation")] 
        public VisioScripting.Models.SelectionOperation Operation { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Shapes !=null)
            {
                this.Client.Selection.Select(this.Shapes);
            }
            else if (this.ShapeIDs!=null)
            {
                this.Client.Selection.Select(this.ShapeIDs);
            }
            else
            {
                if (this.Operation == VisioScripting.Models.SelectionOperation.All)
                {
                    this.Client.Selection.SelectAll();
                }
                else if (this.Operation == VisioScripting.Models.SelectionOperation.None)
                {
                    this.Client.Selection.SelectNone();
                }
                else if (this.Operation == VisioScripting.Models.SelectionOperation.Invert)
                {
                    this.Client.Selection.Invert();
                }
            }
        }
    }
}