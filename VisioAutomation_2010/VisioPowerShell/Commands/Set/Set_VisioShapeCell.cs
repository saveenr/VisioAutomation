using System.Collections;
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
            var writer = new FormulaWriterSIDSRC();
            writer.BlastGuards = this.BlastGuards;
            writer.TestCircular = this.TestCircular;

            var cellmap = VisioAutomation.Scripting.ShapeSheet.CellSRCDictionary.GetCellMapForShapes();
            var valuemap = new VisioAutomation.Scripting.ShapeSheet.CellValueDictionary(cellmap, this.Hashtable);

            var target_shapes = this.Shapes ?? this.Client.Selection.GetShapes();

            this.DumpValues(valuemap);

            foreach (var shape in target_shapes)
            {
                var id = shape.ID16;

                foreach (var cellname in valuemap.CellNames)
                {
                    string cell_value = valuemap[cellname];
                    var cell_src = valuemap.GetSRC(cellname);
                    writer.SetFormula(id,cell_src, cell_value);
                }
            }

            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();

            this.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
            this.WriteVerbose("TestCircular: {0}", this.TestCircular);
            this.WriteVerbose("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerbose("Number of Total Updates: {0}", writer.Count);

            using (var undoscope = this.Client.Application.NewUndoScope( "SetShapeCells"))
            {
                this.WriteVerbose("Start Update");
                writer.Commit(surface);
                this.WriteVerbose("End Update");
            }
        }


    }
}