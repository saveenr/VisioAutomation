using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheetQuery.Results;
using VAQUERY = VisioAutomation.ShapeSheetQuery;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class ShapeSheetCommands : CommandSet
    {
        internal ShapeSheetCommands(Client client) :
            base(client)
        {

        }

        public void SetName(IList<IVisio.Shape> target_shapes, IList<string> names)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (names == null || names.Count < 1)
            {
                // do nothing
                return;
            }

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Text"))
            {
                int numnames = names.Count;

                int up_to = System.Math.Min(numnames, shapes.Count);

                for (int i = 0; i < up_to; i++)
                {
                    var new_name = names[i];

                    if (new_name != null)
                    {
                        var shape = shapes[i];
                        shape.Name = new_name;
                    }
                }
            }
        }


        public VAQUERY.QuerySurface GetShapeSheetSurface()
        {
            var ds = this._client.Draw.GetDrawingSurface();
            var ss = new VAQUERY.QuerySurface(ds.Target);
            return ss;
        }


        public ListResult<T> QueryResults<T>(IList<IVisio.Shape> target_shapes, IList<ShapeSheet.SRC> srcs)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            var shapeids = shapes.Select(s => s.ID).ToList();

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

        public ListResult<string> QueryFormulas(IList<IVisio.Shape> target_shapes, IList<ShapeSheet.SRC> srcs)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();

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

        public ListResult<T> QueryResults<T>(IList<IVisio.Shape> target_shapes, IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();

            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            var query = new VAQUERY.Query();
            var sec = query.AddSection(section);

            int ci = 0;
            foreach (var cell in cells)
            {
                string name = string.Format("Cell{0}", ci);
                sec.AddCell((short)cell, name);
                ci++;
            }

           var results = query.GetResults<T>(surface, shapeids);
            return results;
        }

        public ListResult<string> QueryFormulas(IList<IVisio.Shape> target_shapes, IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();

            var surface = this._client.ShapeSheet.GetShapeSheetSurface();

            var query = new VAQUERY.Query();
            var sec = query.AddSection(section);

            int ci = 0;
            foreach (var cell in cells)
            {
                string name = string.Format("Cell{0}", ci);
                sec.AddCell((short)cell, name);
                ci++;
            }

            var formulas = query.GetFormulas(surface, shapeids);
            return formulas;
        }
        
        public void SetFormula(
            IList<IVisio.Shape> target_shapes, 
            IList<ShapeSheet.SRC> srcs, 
            IList<string> formulas,
            IVisio.VisGetSetArgs flags)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
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


            var shapeids = shapes.Select(s=>s.ID).ToList();
            int num_formulas = formulas.Count;

            var update = new ShapeSheet.Update(shapes.Count*num_formulas);
            update.BlastGuards = ((short)flags & (short)IVisio.VisGetSetArgs.visSetBlastGuards) != 0;
            update.TestCircular = ((short)flags & (short)IVisio.VisGetSetArgs.visSetTestCircular) != 0;

            foreach (var shapeid in shapeids)
            {
                for (int i=0; i<num_formulas;i++)
                {
                    var src = srcs[i];
                    var formula = formulas[i];
                    update.SetFormula((short) shapeid, src, formula);        
                }

            }
            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Set ShapeSheet Formulas"))
            {
                update.Execute(surface);
            }
        }

        public void SetResult(
                IList<IVisio.Shape> target_shapes, 
                IList<ShapeSheet.SRC> srcs,
                IList<string> results, IVisio.VisGetSetArgs flags)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
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

            var shapeids = shapes.Select(s => s.ID).ToList();

            int num_results = results.Count;
            var update = new ShapeSheet.Update(shapes.Count * num_results);
            update.BlastGuards = ((short)flags & (short)IVisio.VisGetSetArgs.visSetBlastGuards) != 0;
            update.TestCircular = ((short)flags & (short)IVisio.VisGetSetArgs.visSetTestCircular) != 0;

            foreach (var shapeid in shapeids)
            {
                for (int i = 0; i < num_results; i++)
                {
                    var src = srcs[i];
                    var result = results[i];
                    update.SetResult((short)shapeid, src, result, IVisio.VisUnitCodes.visNumber);
                }
            }

            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Set ShapeSheet Result"))
            {
                update.Execute(surface);
            }
        }
        
        public void Update(ShapeSheetUpdate update, bool blastguards, bool testcircular)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            this._client.WriteVerbose( "Staring ShapeSheet Update");
            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Update ShapeSheet Formulas"))
            {
                var internal_update = update.update;
                internal_update.BlastGuards = blastguards;
                internal_update.TestCircular = testcircular;
                this._client.WriteVerbose( "BlastGuards={0}", blastguards);
                this._client.WriteVerbose( "TestCircular={0}", testcircular);
                internal_update.Execute(surface);                
            }
            this._client.WriteVerbose( "Ending ShapeSheet Update");
        }
    }
}