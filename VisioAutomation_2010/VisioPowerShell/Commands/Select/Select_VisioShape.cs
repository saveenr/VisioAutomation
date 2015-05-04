using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Select, "VisioShape")]
    public class Select_VisioShape : VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapes")]
        public IVisio.Shape[] Shapes { get; set; }
        
        [SMA.ParameterAttribute(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapeIDs")]
        public int[] ShapeIDs { get; set; }
        
        [SMA.ParameterAttribute(Mandatory = true, Position=0, ParameterSetName = "SelectByOperation")] 
        public SelectionOperation Operation { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Shapes !=null)
            {
                this.client.Selection.Select(this.Shapes);
            }
            else if (this.ShapeIDs!=null)
            {
                this.client.Selection.Select(this.ShapeIDs);
            }
            else
            {
                if (this.Operation == SelectionOperation.All)
                {
                    this.client.Selection.All();
                }
                else if (this.Operation == SelectionOperation.None)
                {
                    this.client.Selection.None();
                }
                else if (this.Operation == SelectionOperation.Invert)
                {
                    this.client.Selection.Invert();
                }
            }
        }
    }
}