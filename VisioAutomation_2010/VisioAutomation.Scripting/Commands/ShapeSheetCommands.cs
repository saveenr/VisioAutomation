using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VAQUERY = VisioAutomation.ShapeSheet.Queries;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class ShapeSheetCommands : CommandSet
    {
        internal ShapeSheetCommands(Client client) :
            base(client)
        {

        }

        internal void __SetCells(TargetShapes targets, VisioAutomation.ShapeSheet.CellGroups.CellGroupBase cells, IVisio.Page page)
        {
            targets = targets.ResolveShapes(this._client);
            var shape_ids = targets.ToShapeIDs();
            var writer = new VisioAutomation.ShapeSheet.ShapeSheetWriter();

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

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(page);
            writer.Commit(surface);
        }

        public void SetName(TargetShapes targets, IList<string> names)
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


        public ShapeSheetSurface GetShapeSheetSurface()
        {
            var drawing_surface = this._client.Draw.GetDrawingSurface();
            var shapesheet_surface = new ShapeSheetSurface(drawing_surface.Target);
            return shapesheet_surface;
        }


        public VisioAutomation.Scripting.ShapeSheet.ShapeSheetWriter GetWriter(IVisio.Page page)
        {
            var writer = new VisioAutomation.Scripting.ShapeSheet.ShapeSheetWriter(this._client, page);
            return writer;
        }

        public VisioAutomation.Scripting.ShapeSheet.ShapeSheetReader GetReader(IVisio.Page page)
        {
            var reader = new VisioAutomation.Scripting.ShapeSheet.ShapeSheetReader(this._client, page);
            return reader;
        }

        public void SetPageCells(TargetShapes targets, Dictionary<string, string> hashtable, bool blast_guards,
            bool test_circular)
        {
            var targets2 = targets.ToShapeIDs();
            this.SetPageCells(targets2,hashtable,blast_guards,test_circular);
        }

        public void SetPageCells(TargetShapeIDs targets, Dictionary<string, string> hashtable, bool blast_guards, bool test_circular)
        {
            var writer = new ShapeSheetWriter();
            writer.BlastGuards = blast_guards;
            writer.TestCircular = test_circular;

            var cellmap = VisioAutomation.Scripting.ShapeSheet.CellSRCDictionary.GetCellMapForPages();
            var valuemap = new VisioAutomation.Scripting.ShapeSheet.CellValueDictionary(cellmap, hashtable);

            foreach (var shape_id in targets.ShapeIDs)
            {
                foreach (var cellname in valuemap.Keys)
                {
                    string cell_value = valuemap[cellname];
                    var cell_src = valuemap.GetSRC(cellname);
                    writer.SetFormula((short)shape_id, cell_src, cell_value);
                }
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

        public void SetShapeCells(TargetShapes targets, Dictionary<string, string> hashtable, bool blast_guards, bool test_circular)
        {
            targets = targets.ResolveShapes(this._client);
            var target_ids= targets.ToShapeIDs();
            this.SetShapeCells(target_ids, hashtable, blast_guards, test_circular);
        }

        public void SetShapeCells(TargetShapeIDs targets, Dictionary<string, string> hashtable, bool blast_guards, bool test_circular)
        {
            var writer = new ShapeSheetWriter();
            writer.BlastGuards = blast_guards;
            writer.TestCircular = test_circular;

            var cellmap = VisioAutomation.Scripting.ShapeSheet.CellSRCDictionary.GetCellMapForShapes();
            var valuemap = new VisioAutomation.Scripting.ShapeSheet.CellValueDictionary(cellmap, hashtable);

            foreach (var shape_id in targets.ShapeIDs)
            {
                foreach (var cellname in valuemap.Keys)
                {
                    string cell_value = valuemap[cellname];
                    var cell_src = valuemap.GetSRC(cellname);
                    writer.SetFormula((short)shape_id, cell_src, cell_value);
                }
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