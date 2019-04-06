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
            var targetshapes = new VisioScripting.TargetShapes(this.Shapes);
            var dicof_shape_to_udcelldic = this.Client.UserDefinedCell.GetUserDefinedCells(targetshapes, CellValueType.Formula);

            if (this.GetCells)
            {
                this.WriteObject(dicof_shape_to_udcelldic);
                return;
            }

            foreach (var shape_udcelldic_pair in dicof_shape_to_udcelldic)
            {
                var shape = shape_udcelldic_pair.Key;
                var udcelldic = shape_udcelldic_pair.Value;
                int shapeid = shape.ID;
                foreach (var udcellname_udcellcells_pair in udcelldic)
                {
                    string udcellname = udcellname_udcellcells_pair.Key;
                    var udcellcells = udcellname_udcellcells_pair.Value;
                    string udcell_value = udcellcells.Value.ToString();
                    string udcell_prompt = udcellcells.Prompt.ToString();

                    var udcell_obj = new VisioPowerShell.Models.UserDefinedCell(shapeid, udcellname, udcell_value, udcell_prompt);
                    this.WriteObject(udcell_obj);
                }
            }

            this.WriteObject(dicof_shape_to_udcelldic);
        }
    }
}