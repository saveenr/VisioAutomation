using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Writers;
using VisioAutomation.ShapeSheet.Queries.Outputs;
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
            var resolved_targets = targets.ResolveShapes(this._client);
            var shape_ids = resolved_targets.ToShapeIDs();
            var writer = new VisioAutomation.ShapeSheet.Writers.FormulaWriterSIDSRC();

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
        public void SetName(TargetShapes targets, IList<string> names)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (names == null || names.Count < 1)
            {
                // do nothing
                return;
            }

            var shapes = targets.ResolveShapes(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Text"))
            {
                int numnames = names.Count;

                int up_to = System.Math.Min(numnames, shapes.Shapes.Count);

                for (int i = 0; i < up_to; i++)
                {
                    var new_name = names[i];

                    if (new_name != null)
                    {
                        var shape = shapes.Shapes[i];
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


        public ListOutput<T> QueryResults<T>(TargetShapes targets, IList<VisioAutomation.ShapeSheet.SRC> srcs)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);

            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            var shapeids = shapes.Shapes.Select(s => s.ID).ToList();

            var query = new VAQUERY.Query();

            int ci = 0;
            foreach (var src in srcs)
            {
                string colname = string.Format("Col{0}", ci);
                query.AddCell(src, colname);
                ci++;
            }

            var results = query.GetResults<T>(surface, shapeids);
            return results;
        }

        public ListOutput<string> QueryFormulas(TargetShapes targets, IList<VisioAutomation.ShapeSheet.SRC> srcs)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);

            var shapeids = shapes.Shapes.Select(s => s.ID).ToList();

            var surface = this._client.ShapeSheet.GetShapeSheetSurface();

            var query = new VAQUERY.Query();

            int ci = 0;
            foreach (var src in srcs)
            {
                string colname = string.Format("Col{0}", ci);
                query.AddCell(src, colname);
                ci++;
            }

            var formulas = query.GetFormulas(surface, shapeids);

            return formulas;
        }

        public ListOutput<T> QueryResults<T>(TargetShapes targets, IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);

            var shapeids = shapes.Shapes.Select(s => s.ID).ToList();

            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            var query = new VAQUERY.Query();
            var sec = query.AddSubQuery(section);

            int ci = 0;
            foreach (var cell in cells)
            {
                string name = string.Format("Cell{0}", ci);
                var src = new SRC(section,0,cell);
                sec.AddCell(src, name);
                ci++;
            }

           var results = query.GetResults<T>(surface, shapeids);
            return results;
        }

        public ListOutput<string> QueryFormulas(TargetShapes targets, IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);

            var shapeids = shapes.Shapes.Select(s => s.ID).ToList();

            var surface = this._client.ShapeSheet.GetShapeSheetSurface();

            var query = new VAQUERY.Query();
            var sec = query.AddSubQuery(section);

            int ci = 0;
            foreach (var cell in cells)
            {
                string name = string.Format("Cell{0}", ci);
                var src = new SRC(section, 0, cell);
                sec.AddCell(src, name);
                ci++;
            }

            var formulas = query.GetFormulas(surface, shapeids);
            return formulas;
        }

        public ShapeSheetWriter GetWriter(IVisio.Page page)
        {
            var w = new ShapeSheetWriter(this._client, page);
            return w;
        }

        public void SetPageCells(TargetShapes targets, System.Collections.Hashtable hashtable, bool blast_guards,
            bool test_circular)
        {
            var targets2 = targets.ToShapeIDs();
            this.SetPageCells(targets2,hashtable,blast_guards,test_circular);
        }

        public void SetPageCells(TargetShapeIDs targets,System.Collections.Hashtable hashtable, bool blast_guards, bool test_circular)
        {
            var writer = new FormulaWriterSIDSRC();
            writer.BlastGuards = blast_guards;
            writer.TestCircular = test_circular;

            var cellmap = VisioAutomation.Scripting.ShapeSheet.CellSRCDictionary.GetCellMapForPages();
            var valuemap = new VisioAutomation.Scripting.ShapeSheet.CellValueDictionary(cellmap, hashtable);

            foreach (var shape_id in targets.ShapeIDs)
            {
                foreach (var cellname in valuemap.CellNames)
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
            this._client.WriteVerbose("Number of Total Updates: {0}", writer.Count);

            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Cells"))
            {
                this._client.WriteVerbose("Start Update");
                writer.Commit(surface);
                this._client.WriteVerbose("End Update");
            }
        }

        public void SetShapeCells(TargetShapes targets, System.Collections.Hashtable hashtable, bool blast_guards, bool test_circular)
        {
            var resolved_targets = targets.ResolveShapes(this._client);
            var target_ids= resolved_targets.ToShapeIDs();
            this.SetShapeCells(target_ids, hashtable, blast_guards, test_circular);
        }

        public void SetShapeCells(TargetShapeIDs targets, System.Collections.Hashtable hashtable, bool blast_guards, bool test_circular)
        {
            var writer = new FormulaWriterSIDSRC();
            writer.BlastGuards = blast_guards;
            writer.TestCircular = test_circular;

            var cellmap = VisioAutomation.Scripting.ShapeSheet.CellSRCDictionary.GetCellMapForShapes();
            var valuemap = new VisioAutomation.Scripting.ShapeSheet.CellValueDictionary(cellmap, hashtable);

            foreach (var shape_id in targets.ShapeIDs)
            {
                foreach (var cellname in valuemap.CellNames)
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
            this._client.WriteVerbose("Number of Total Updates: {0}", writer.Count);

            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Cells"))
            {
                this._client.WriteVerbose("Start Update");
                writer.Commit(surface);
                this._client.WriteVerbose("End Update");
            }
        }
    }
}