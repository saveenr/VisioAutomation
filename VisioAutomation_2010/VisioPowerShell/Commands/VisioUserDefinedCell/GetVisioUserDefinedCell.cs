using VASS=VisioAutomation.ShapeSheet;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioUserDefinedCell
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioUserDefinedCell)]
    public class GetVisioUserDefinedCell : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shapes);
            var type = VASS.CellValueType.Formula;
            var dicof_shape_to_udcelldic = this.Client.UserDefinedCell.GetUserDefinedCellsAsShapeDictionary(targetshapes, type);

            this.WriteObject(dicof_shape_to_udcelldic);
        }
    }
}