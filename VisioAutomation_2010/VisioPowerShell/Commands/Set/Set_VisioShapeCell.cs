using System.Collections;
using System.Linq;
using System.Management.Automation;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Set
{
    [Cmdlet(VerbsCommon.Set, VisioPowerShell.Nouns.VisioShapeCell)]
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
            var update = new FormulaWriterSIDSRC();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular = this.TestCircular;

            var cellmap = CellSRCDictionary.GetCellMapForShapes();
            var valuemap = new CellValueDictionary(cellmap, this.Hashtable);

            var target_shapes = this.Shapes ?? this.Client.Selection.GetShapes();

            this.DumpValues(valuemap);

            foreach (var shape in target_shapes)
            {
                var id = shape.ID16;

                foreach (var cellname in valuemap.CellNames)
                {
                    string cell_value = valuemap[cellname];
                    var cell_src = valuemap.GetSRC(cellname);
                    update.SetFormula(id,cell_src, cell_value);
                }
            }

            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();

            this.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
            this.WriteVerbose("TestCircular: {0}", this.TestCircular);
            this.WriteVerbose("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerbose("Number of Total Updates: {0}", update.Count);

            using (var undoscope = this.Client.Application.NewUndoScope( "SetShapeCells"))
            {
                this.WriteVerbose("Start Update");
                update.Commit(surface);
                this.WriteVerbose("End Update");
            }
        }


    }
}