using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class ShapeSheetCommands : CommandSet
    {
        internal ShapeSheetCommands(Client client) :
            base(client)
        {

        }

        internal void __SetCells(VisioScripting.Models.TargetShapes targets, VisioAutomation.ShapeSheet.CellGroups.CellGroupBase cells, IVisio.Page page)
        {
            targets = targets.ResolveShapes(this._client);
            var shape_ids = targets.ToShapeIDs();
            var writer = new SidSrcWriter();

            foreach (var shape_id in shape_ids.ShapeIDs)
            {
                if (cells is VisioAutomation.ShapeSheet.CellGroups.CellGroupMultiRow)
                {
                    var cells_mr = (VisioAutomation.ShapeSheet.CellGroups.CellGroupMultiRow)cells;
                    cells_mr.SetFormulas((short)shape_id, writer, 0);
                }
                else
                {
                    var cells_sr = (VisioAutomation.ShapeSheet.CellGroups.CellGroupSingleRow)cells;
                    cells_sr.SetFormulas((short)shape_id, writer);
                }
            }

            writer.Commit(page);
        }

        public void SetName(VisioScripting.Models.TargetShapes targets, IList<string> names)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (names == null || names.Count < 1)
            {
                // do nothing
                return;
            }

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Text"))
            {
                int numnames = names.Count;

                int up_to = System.Math.Min(numnames, targets.Shapes.Count);

                for (int i = 0; i < up_to; i++)
                {
                    var new_name = names[i];

                    if (new_name != null)
                    {
                        var shape = targets.Shapes[i];
                        shape.Name = new_name;
                    }
                }
            }
        }


        public VisioAutomation.SurfaceTarget GetShapeSheetSurface()
        {
            var drawing_surface = this._client.Draw.GetDrawingSurface();
            var shapesheet_surface = drawing_surface;
            return shapesheet_surface;
        }


        public VisioScripting.Models.ShapeSheetWriter GetWriter(IVisio.Page page)
        {
            var writer = new VisioScripting.Models.ShapeSheetWriter(this._client, page);
            return writer;
        }

        public VisioScripting.Models.ShapeSheetReader GetReader(IVisio.Page page)
        {
            var reader = new VisioScripting.Models.ShapeSheetReader(this._client, page);
            return reader;
        }

        public void SetPageCells(VisioScripting.Models.TargetShapes targets, VisioScripting.Models.BaseCells cells, bool blast_guards,
            bool test_circular)
        {
            var targets2 = targets.ToShapeIDs();
            this.SetPageCells(targets2,cells,blast_guards,test_circular);
        }

        public void SetPageCells(VisioScripting.Models.TargetShapeIDs targets, VisioScripting.Models.BaseCells cells, bool blast_guards, bool test_circular)
        {
            var writer = new SidSrcWriter();
            writer.BlastGuards = blast_guards;
            writer.TestCircular = test_circular;


            foreach (var shape_id in targets.ShapeIDs)
            {
                cells.Apply(writer, (short) shape_id);
            }

            var surface = this._client.ShapeSheet.GetShapeSheetSurface();

            this._client.WriteVerbose("BlastGuards: {0}", blast_guards);
            this._client.WriteVerbose("TestCircular: {0}", test_circular);
            this._client.WriteVerbose("Number of Shapes : {0}", targets.ShapeIDs.Count);

            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Cells"))
            {
                this._client.WriteVerbose("Start Update");
                writer.Commit(surface);
                this._client.WriteVerbose("End Update");
            }
        }

        public void SetShapeCells(VisioScripting.Models.TargetShapes targets, VisioScripting.Models.BaseCells cells, bool blast_guards, bool test_circular)
        {
            targets = targets.ResolveShapes(this._client);
            var target_ids = targets.ToShapeIDs();
            this.SetShapeCells(target_ids, cells, blast_guards, test_circular);
        }

        public void SetShapeCells(VisioScripting.Models.TargetShapeIDs targets, VisioScripting.Models.BaseCells cells, bool blast_guards, bool test_circular)
        {
            var writer = new SidSrcWriter();
            writer.BlastGuards = blast_guards;
            writer.TestCircular = test_circular;


            foreach (var shape_id in targets.ShapeIDs)
            {
                cells.Apply(writer, (short)shape_id);
            }

            var surface = this._client.ShapeSheet.GetShapeSheetSurface();

            this._client.WriteVerbose("BlastGuards: {0}", blast_guards);
            this._client.WriteVerbose("TestCircular: {0}", test_circular);
            this._client.WriteVerbose("Number of Shapes : {0}", targets.ShapeIDs.Count);

            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Cells"))
            {
                this._client.WriteVerbose("Start Update");
                writer.Commit(surface);
                this._client.WriteVerbose("End Update");
            }
        }

    }
}