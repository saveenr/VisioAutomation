using VisioPowerShell.Models;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Select, Nouns.VisioShape)]
    public class SelectVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByShapes")]
        public IVisio.Shape[] Shapes { get; set; }

        [SMA.Parameter(Mandatory = true, Position = 0, ParameterSetName = "SelectByOperation")]
        public VisioPowerShell.Models.ShapeSelectionOperation SelectionOperation { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Shapes !=null)
            {
                this.Client.Selection.SelectShapes(VisioScripting.TargetWindow.Auto, this.Shapes);
                return;
            }

            if (this.SelectionOperation == ShapeSelectionOperation.SelectAll)
            {
                this.Client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);
            }
            else if (this.SelectionOperation == ShapeSelectionOperation.SelectNone)
            {
                this.Client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            }
            else if (this.SelectionOperation == ShapeSelectionOperation.InvertSelection)
            {
                this.Client.Selection.InvertSelection(VisioScripting.TargetWindow.Auto);
            }
        }
    }
}