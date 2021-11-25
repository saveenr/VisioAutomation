

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Select, Nouns.VisioShape)]
    public class SelectVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapes")]
        public IVisio.Shape[] Shapes { get; set; }

        [SMA.Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByOperation")]
        public VisioScripting.Models.ShapeSelectionOperation SelectionOperation { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Shapes !=null)
            {
                this.Client.Selection.SelectShapes(VisioScripting.TargetWindow.Auto, this.Shapes);
                return;
            }

            this.Client.Selection.SelectShapeOperation(VisioScripting.TargetWindow.Auto, this.SelectionOperation);
        }
    }
}