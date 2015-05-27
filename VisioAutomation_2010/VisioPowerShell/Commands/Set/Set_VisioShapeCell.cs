using System.Collections;
using System.Linq;
using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Set
{
    [Cmdlet(VerbsCommon.Set, "VisioShapeCell")]
    public class Set_VisioShapeCell : VisioCmdlet
    {
        [Parameter(Mandatory = false, Position = 0)]
        public Hashtable Hashtable { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter BlastGuards { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter TestCircular { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        protected override void ProcessRecord()
        {
            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular = this.TestCircular;

            var cellmap = CellSRCDictionary.GetCellMapForShapes();
            var valuemap = new CellValueDictionary(cellmap, this.Hashtable);

            var target_shapes = this.Shapes ?? this.client.Selection.GetShapes();

            this.DumpValues(valuemap);

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

            using (var undoscope = this.client.Application.NewUndoScope( "SetShapeCells"))
            {
                this.WriteVerbose("Start Update");
                update.Execute(surface);
                this.WriteVerbose("End Update");
            }
        }


    }
}