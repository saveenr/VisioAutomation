using VisioAutomation.ShapeSheet;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioUserDefinedCell
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioUserDefinedCell)]
    public class GetVisioUserDefinedCell : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            var dicof_shape_to_udcelldic = this.Client.UserDefinedCell.GetUserDefinedCells(targets, CellValueType.Formula);

            if (this.GetCells)
            {
                this.WriteObject(dicof_shape_to_udcelldic);
                return;
            }

            this.WriteObject(dicof_shape_to_udcelldic);
        }
    }
}