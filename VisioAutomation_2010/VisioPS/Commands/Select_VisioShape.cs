using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Select, "VisioShape")]
    public class Select_VisioShape : VisioPS.VisioPSCmdlet
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
                if (this.Operation == VisioPS.SelectionOperation.All)
                {
                    scriptingsession.Selection.SelectAll();
                }
                else if (this.Operation == VisioPS.SelectionOperation.None)
                {
                    scriptingsession.Selection.SelectNone();
                }
                else if (this.Operation == VisioPS.SelectionOperation.Invert)
                {
                    scriptingsession.Selection.SelectInvert();
                }
            }
        }
    }
}