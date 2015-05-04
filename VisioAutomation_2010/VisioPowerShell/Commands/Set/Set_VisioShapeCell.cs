using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA = VisioAutomation;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioShapeCell")]
    public class Set_VisioShapeCell : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position = 0)]
        public System.Collections.Hashtable Hashtable { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        protected override void ProcessRecord()
        {
            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular = this.TestCircular;

            var cellmap = CellMap.GetCellMapForShapes();
            var valuemap = new CellValueMap(cellmap);
            valuemap.UpdateValueMap(this.Hashtable);

            var target_shapes = this.Shapes ?? this.client.Selection.GetShapes();

            foreach (var shape in target_shapes)
            {
                var id = shape.ID16;

                foreach (var cellname in valuemap.CellNames)
                {
                    string cell_value = valuemap[cellname];
                    var cell_src = valuemap.GetSRC(cellname);
                    update.SetFormulaIgnoreNull(id,cell_src, cell_value);
                }
            }

            var surface = this.client.ShapeSheet.GetShapeSheetSurface();

            this.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
            this.WriteVerbose("TestCircular: {0}", this.TestCircular);
            this.WriteVerbose("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerbose("Number of Total Updates: {0}", update.Count());

            using (var undoscope = new VA.Application.UndoScope(this.client.VisioApplication, "SetShapeCells"))
            {
                this.WriteVerbose("Start Update");
                update.Execute(surface);
                this.WriteVerbose("End Update");
            }
        }
    }
}