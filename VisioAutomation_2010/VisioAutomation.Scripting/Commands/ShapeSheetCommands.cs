using System.Collections.Generic;
using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ShapeSheetCommands : CommandSet
    {
        public ShapeSheetCommands(Session session) :
            base(session)
        {

        }

        public VA.ShapeSheet.Query.CellQuery.QueryResultList<T> QueryResults<T>(IList<IVisio.Shape> target_shapes, IList<VA.ShapeSheet.SRC> srcs)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var surface = this.Session.Draw.GetDrawingSurfaceSafe();
            var shapeids = shapes.Select(s => s.ID).ToList();

            var query = new VA.ShapeSheet.Query.CellQuery();

            int ci = 0;
            foreach (var src in srcs)
            {
                string colname = string.Format("Col{0}", ci);
                query.Columns.Add(src,colname);
                ci++;
            }

            var results = query.GetResults<T>(surface, shapeids);
            return results;
        }

        public VA.ShapeSheet.Query.CellQuery.QueryResultList<string> QueryFormulas(IList<IVisio.Shape> target_shapes, IList<VA.ShapeSheet.SRC> srcs)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();

            var surface = this.Session.Draw.GetDrawingSurfaceSafe();
 
            var query = new VA.ShapeSheet.Query.CellQuery();

            int ci = 0;
            foreach (var src in srcs)
            {
                string colname = string.Format("Col{0}", ci);
                query.Columns.Add(src,colname);
                ci++;
            }

            var formulas = query.GetFormulas(surface, shapeids);

            return formulas;
        }

        public VA.ShapeSheet.Query.CellQuery.QueryResultList<T> QueryResults<T>(IList<IVisio.Shape> target_shapes, IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();

            var app = this.Session.VisioApplication;
            var surface = this.Session.Draw.GetDrawingSurfaceSafe();
            var query = new VA.ShapeSheet.Query.CellQuery();
            var sec = query.Sections.Add(section);

            int ci = 0;
            foreach (var cell in cells)
            {
                string name = string.Format("Cell{0}", ci);
                sec.Columns.Add((short)cell, name);
                ci++;
            }

           var results = query.GetResults<T>(surface, shapeids);
            return results;
        }

        public VA.ShapeSheet.Query.CellQuery.QueryResultList<string> QueryFormulas(IList<IVisio.Shape> target_shapes, IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();

            var surface = this.Session.Draw.GetDrawingSurfaceSafe();

            var query = new VA.ShapeSheet.Query.CellQuery();
            var sec = query.Sections.Add(section);

            int ci = 0;
            foreach (var cell in cells)
            {
                string name = string.Format("Cell{0}", ci);
                sec.Columns.Add((short)cell, name);
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
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                this.Session.WriteVerbose("SetFormula: Zero Shapes. Not performing Operation");
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
                this.Session.WriteVerbose("SetFormula: One of the Input Formulas is a NULL value");
                throw new System.ArgumentException("formulas contains a null value");
            }

            this.Session.WriteVerbose("SetFormula: src count= {0} and formula count = {1}", srcs.Count, formulas.Count);

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
            var surface = this.Session.Draw.GetDrawingSurfaceSafe();
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Set ShapeSheet Formulas"))
            {
                update.Execute(surface);
            }
        }

        public void SetResult(
                IList<IVisio.Shape> target_shapes, 
                IList<VA.ShapeSheet.SRC> srcs,
                IList<string> results, IVisio.VisGetSetArgs flags)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                this.Session.WriteVerbose("SetResult: Zero Shapes. Not performing Operation");
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
                this.Session.WriteVerbose("SetResult: One of the Input Results is a NULL value");
                throw new System.ArgumentException("results contains a null value");
            }

            this.Session.WriteVerbose("SetResult: src count= {0} and result count = {1}", srcs.Count, results.Count);

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

            var surface = this.Session.Draw.GetDrawingSurfaceSafe();
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, "Set ShapeSheet Result"))
            {
                update.Execute(surface);
            }
        }
        
        public void Update(ShapeSheetUpdate update, bool blastguards, bool testcircular)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            this.Session.WriteVerbose( "Staring ShapeSheet Update");
            var surface = this.Session.Draw.GetDrawingSurfaceSafe();
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, "Update ShapeSheet Formulas"))
            {
                var internal_update = update.update;
                internal_update.BlastGuards = blastguards;
                internal_update.TestCircular = testcircular;
                this.Session.WriteVerbose( "BlastGuards={0}", blastguards);
                this.Session.WriteVerbose( "TestCircular={0}", testcircular);
                internal_update.Execute(surface);                
            }
            this.Session.WriteVerbose( "Ending ShapeSheet Update");
        }
    }
}