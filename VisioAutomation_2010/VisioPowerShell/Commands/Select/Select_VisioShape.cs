using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Select, "VisioShape")]
    public class Select_VisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapes")]
        public IVisio.Shape[] Shapes { get; set; }
        
        [SMA.Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapeIDs")]
        public int[] ShapeIDs { get; set; }
        
        [SMA.Parameter(Mandatory = true, Position=0, ParameterSetName = "SelectByOperation")] 
        public SelectionOperation Operation { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            if ( Shapes !=null)
            {
                scriptingsession.Selection.Select(Shapes);
            }
            else if (ShapeIDs!=null)
            {
                scriptingsession.Selection.Select(ShapeIDs);
            }
            else
            {
                if (this.Operation == SelectionOperation.All)
                {
                    scriptingsession.Selection.All();
                }
                else if (this.Operation == SelectionOperation.None)
                {
                    scriptingsession.Selection.None();
                }
                else if (this.Operation == SelectionOperation.Invert)
                {
                    scriptingsession.Selection.Invert();
                }
            }
        }
    }
}