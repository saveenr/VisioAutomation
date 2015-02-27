using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet.Query;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ShapeSheetCommands : CommandSet
    {
        public ShapeSheetCommands(Client client) :
            base(client)
        {

        }

        public VA.ShapeSheet.ShapeSheetSurface GetShapeSheetSurface()
        {
            var ds = this.Client.Draw.GetDrawingSurface();
            var ss = new VA.ShapeSheet.ShapeSheetSurface(ds.Target);
            return ss;
        }


        public QueryResultList<T> QueryResults<T>(IList<IVisio.Shape> target_shapes, IList<VA.ShapeSheet.SRC> srcs)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            var shapeids = shapes.Select(s => s.ID).ToList();

            var query = new VA.ShapeSheet.Query.CellQuery();

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

        public QueryResultList<string> QueryFormulas(IList<IVisio.Shape> target_shapes, IList<VA.ShapeSheet.SRC> srcs)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();

            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
 
            var query = new VA.ShapeSheet.Query.CellQuery();

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

        public QueryResultList<T> QueryResults<T>(IList<IVisio.Shape> target_shapes, IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();
            var app = this.Client.VisioApplication;
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            var query = new VA.ShapeSheet.Query.CellQuery();
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

        public QueryResultList<string> QueryFormulas(IList<IVisio.Shape> target_shapes, IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();

            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();

            var query = new VA.ShapeSheet.Query.CellQuery();
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
            IList<VA.ShapeSheet.SRC> srcs, 
            IList<string> formulas,
            IVisio.VisGetSetArgs flags)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                this.Client.WriteVerbose("SetFormula: Zero Shapes. Not performing Operation");
                return;
            }

            if (srcs == null)
            {
                throw new System.ArgumentNullException("srcs");
            }

            if (formulas == null)
            {
                throw new System.ArgumentNullException("formulas");
            }

            if (formulas.Any( f => f == null))
            {
                this.Client.WriteVerbose("SetFormula: One of the Input Formulas is a NULL value");
                throw new System.ArgumentException("formulas contains a null value");
            }

            this.Client.WriteVerbose("SetFormula: src count= {0} and formula count = {1}", srcs.Count, formulas.Count);

            if (formulas.Count != srcs.Count)
            {
                string msg = string.Format("SetFormula: Must have the same number of srcs ({0}) and formulas ({1})", srcs.Count,formulas.Count);
                throw new System.ArgumentException(msg);
            }


            var shapeids = shapes.Select(s=>s.ID).ToList();
            int num_formulas = formulas.Count;

            var update = new VA.ShapeSheet.Update(shapes.Count*num_formulas);
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
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Set ShapeSheet Formulas"))
            {
                update.Execute(surface);
            }
        }

        public void SetResult(
                IList<IVisio.Shape> target_shapes, 
                IList<VA.ShapeSheet.SRC> srcs,
                IList<string> results, IVisio.VisGetSetArgs flags)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                this.Client.WriteVerbose("SetResult: Zero Shapes. Not performing Operation");
                return;
            }

            if (srcs == null)
            {
                throw new System.ArgumentNullException("srcs");
            }

            if (results == null)
            {
                throw new System.ArgumentNullException("results");
            }

            if (results.Any(f => f == null))
            {
                this.Client.WriteVerbose("SetResult: One of the Input Results is a NULL value");
                throw new System.ArgumentException("results contains a null value");
            }

            this.Client.WriteVerbose("SetResult: src count= {0} and result count = {1}", srcs.Count, results.Count);

            if (results.Count != srcs.Count)
            {
                string msg = string.Format("Must have the same number of srcs ({0}) and results ({1})", srcs.Count, results.Count);
                throw new System.ArgumentException(msg);
            }

            var shapeids = shapes.Select(s => s.ID).ToList();

            int num_results = results.Count;
            var update = new VA.ShapeSheet.Update(shapes.Count * num_results);
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

            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Set ShapeSheet Result"))
            {
                update.Execute(surface);
            }
        }
        
        public void Update(ShapeSheetUpdate update, bool blastguards, bool testcircular)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            this.Client.WriteVerbose( "Staring ShapeSheet Update");
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Update ShapeSheet Formulas"))
            {
                var internal_update = update.update;
                internal_update.BlastGuards = blastguards;
                internal_update.TestCircular = testcircular;
                this.Client.WriteVerbose( "BlastGuards={0}", blastguards);
                this.Client.WriteVerbose( "TestCircular={0}", testcircular);
                internal_update.Execute(surface);                
            }
            this.Client.WriteVerbose( "Ending ShapeSheet Update");
        }
    }
}