using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Select
{
    [Cmdlet(VerbsCommon.Select, VisioPowerShell.Nouns.VisioShape)]
    public class Select_VisioShape : VisioCmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapes")]
        public IVisio.Shape[] Shapes { get; set; }
        
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapeIDs")]
        public int[] ShapeIDs { get; set; }
        
        [Parameter(Mandatory = true, Position=0, ParameterSetName = "SelectByOperation")] 
        public Model.SelectionOperation Operation { get; set; }

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
                if (this.Operation == Model.SelectionOperation.All)
                {
                    this.Client.Selection.SelectAll();
                }
                else if (this.Operation == Model.SelectionOperation.None)
                {
                    this.Client.Selection.SelectNone();
                }
                else if (this.Operation == Model.SelectionOperation.Invert)
                {
                    this.Client.Selection.Invert();
                }
            }
        }
    }
}