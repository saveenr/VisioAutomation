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



        public void SetFormula(
            TargetShapes targets, 
            IList<VisioAutomation.ShapeSheet.SRC> srcs, 
            IList<string> formulas,
            IVisio.VisGetSetArgs flags)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);

            if (shapes.Shapes.Count < 1)
            {
                this._client.WriteVerbose("SetFormula: Zero Shapes. Not performing Operation");
                return;
            }

            if (srcs == null)
            {
                throw new System.ArgumentNullException(nameof(srcs));
            }

            if (formulas == null)
            {
                throw new System.ArgumentNullException(nameof(formulas));
            }

            if (formulas.Any( f => f == null))
            {
                this._client.WriteVerbose("SetFormula: One of the Input Formulas is a NULL value");
                throw new System.ArgumentException("Formulas contains a null value");
            }

            this._client.WriteVerbose("SetFormula: src count= {0} and formula count = {1}", srcs.Count, formulas.Count);

            if (formulas.Count != srcs.Count)
            {
                string msg =
                    string.Format("SetFormula: Must have the same number of srcs ({0}) and formulas ({1})", srcs.Count,
                        formulas.Count);
                throw new System.ArgumentException(msg, nameof(formulas));
            }


            var shapeids = shapes.Shapes.Select(s=>s.ID).ToList();
            int num_formulas = formulas.Count;

            var writer = new FormulaWriterSIDSRC(shapes.Shapes.Count*num_formulas);
            writer.BlastGuards = ((short)flags & (short)IVisio.VisGetSetArgs.visSetBlastGuards) != 0;
            writer.TestCircular = ((short)flags & (short)IVisio.VisGetSetArgs.visSetTestCircular) != 0;

            foreach (var shapeid in shapeids)
            {
                for (int i=0; i<num_formulas;i++)
                {
                    var src = srcs[i];
                    var formula = formulas[i];
                    writer.SetFormula((short) shapeid, src, formula);        
                }

            }
            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            using (var undoscope = this._client.Application.NewUndoScope("Set ShapeSheet Formulas"))
            {
                writer.Commit(surface);
            }
        }

        public void SetResult<T>(
                TargetShapes  targets, 
                IList<VisioAutomation.ShapeSheet.SRC> srcs,
                IList<string> results, IVisio.VisGetSetArgs flags)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes(this._client);

            if (shapes.Shapes.Count < 1)
            {
                this._client.WriteVerbose("SetResult: Zero Shapes. Not performing Operation");
                return;
            }

            if (srcs == null)
            {
                throw new System.ArgumentNullException(nameof(srcs));
            }

            if (results == null)
            {
                throw new System.ArgumentNullException(nameof(results));
            }

            if (results.Any(f => f == null))
            {
                this._client.WriteVerbose("SetResult: One of the Input Results is a NULL value");
                throw new System.ArgumentException("results contains a null value",nameof(results));
            }

            this._client.WriteVerbose("SetResult: src count= {0} and result count = {1}", srcs.Count, results.Count);

            if (results.Count != srcs.Count)
            {
                string msg = string.Format("Must have the same number of srcs ({0}) and results ({1})", srcs.Count,
                    results.Count);
                throw new System.ArgumentException(msg,nameof(results));
            }

            var shapeids = shapes.Shapes.Select(s => s.ID).ToList();

            int num_results = results.Count;
            var writer = new ResultWriterSIDSRC(shapes.Shapes.Count * num_results);
            writer.BlastGuards = ((short)flags & (short)IVisio.VisGetSetArgs.visSetBlastGuards) != 0;
            writer.TestCircular = ((short)flags & (short)IVisio.VisGetSetArgs.visSetTestCircular) != 0;

            foreach (var shapeid in shapeids)
            {
                for (int i = 0; i < num_results; i++)
                {
                    var result = results[i];
                    var streamitem = new SIDSRC((short) shapeid, srcs[i]);
                    writer.SetResult(streamitem, result, IVisio.VisUnitCodes.visNumber);
                }
            }

            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            using (var undoscope = this._client.Application.NewUndoScope("Set ShapeSheet Result"))
            {
                writer.Commit(surface);
            }
        }
        
        public void Commit(ShapeSheetWriter writer, bool blastguards, bool testcircular)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            using (var undoscope = this._client.Application.NewUndoScope("Modify ShapeSheet"))
            {
                var internal_writer = writer.formula_writer;
                internal_writer.BlastGuards = blastguards;
                internal_writer.TestCircular = testcircular;
                this._client.WriteVerbose( "BlastGuards={0}", blastguards);
                this._client.WriteVerbose( "TestCircular={0}", testcircular);
                internal_writer.Commit(surface);                
            }
        }

        public void SetPageCells(TargetShapes targets, System.Collections.Hashtable ht, bool blast_guards,
            bool test_circular)
        {
            var page = this._client.Page.Get();
            var targets2 = targets.ToShapeIDs( page );
            this.SetPageCells(targets2,ht,blast_guards,test_circular);
        }

        public void SetPageCells(TargetShapeIDs targets,System.Collections.Hashtable ht, bool blast_guards, bool test_circular)
        {
            var writer = new FormulaWriterSIDSRC();
            writer.BlastGuards = blast_guards;
            writer.TestCircular = test_circular;

            var cellmap = VisioAutomation.Scripting.ShapeSheet.CellSRCDictionary.GetCellMapForPages();
            var valuemap = new VisioAutomation.Scripting.ShapeSheet.CellValueDictionary(cellmap, ht);

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
            var page = this._client.Page.Get();
            var resolved_targets = targets.ResolveShapes(this._client);
            var target_ids= resolved_targets.ToShapeIDs(page);
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