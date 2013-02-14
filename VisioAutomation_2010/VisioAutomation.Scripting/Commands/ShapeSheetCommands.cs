using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
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

        public VA.ShapeSheet.Data.Table<T> QueryResults<T>( IList<IVisio.Shape> target_shapes, IList<VA.ShapeSheet.SRC> srcs)
        {
            var shapes = this.get_target_shapes(target_shapes);
            var app = this.Session.VisioApplication;
            var page = app.ActivePage;
            var shapeids = shapes.Select(s=>s.ID).ToList();

            var query = new VA.ShapeSheet.Query.CellQuery();

            int ci = 0;
            foreach (var src in srcs)
            {
                query.AddColumn(src);
                ci++;
            }

            var results = query.GetResults<T>(page, shapeids);
            return results;
        }

        public VA.ShapeSheet.Data.Table<string> QueryFormulas(IList<IVisio.Shape> target_shapes, IList<VA.ShapeSheet.SRC> srcs)
        {
            var shapes = this.get_target_shapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();

            var app = this.Session.VisioApplication;
            var page = app.ActivePage;
            var active_window = app.ActiveWindow;
            
            var query = new VA.ShapeSheet.Query.CellQuery();

            int ci = 0;
            foreach (var src in srcs)
            {
                query.AddColumn(src);
                ci++;
            }

            var formulas = query.GetFormulas(page, shapeids);

            return formulas;
        }

        public VA.ShapeSheet.Data.Table<T> QueryResults<T>(IList<IVisio.Shape> target_shapes, IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            var shapes = this.get_target_shapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();

            var app = this.Session.VisioApplication;
            var page = app.ActivePage;
            var query = new VA.ShapeSheet.Query.SectionQuery((short)section);

            int ci = 0;
            foreach (var cell in cells)
            {
                query.AddColumn(cell);
                ci++;
            }

            var results = query.GetResults<T>(page, shapeids);
            return results;
        }

        public VA.ShapeSheet.Data.Table<string> QueryFormulas(IList<IVisio.Shape> target_shapes, IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            var shapes = this.get_target_shapes(target_shapes);
            var shapeids = shapes.Select(s => s.ID).ToList();

            var app = this.Session.VisioApplication;
            var page = app.ActivePage;

            var query = new VA.ShapeSheet.Query.SectionQuery((short)section);

            int ci = 0;
            foreach (var cell in cells)
            {
                query.AddColumn(cell);
                ci++;
            }

            var formulas = query.GetFormulas(page, shapeids);
            return formulas;
        }
        
        public void SetFormula(IList<IVisio.Shape> target_shapes, IList<VA.ShapeSheet.SRC> srcs, 
            IList<string> formulas,
            IVisio.VisGetSetArgs flags)
        {
            var shapes = this.get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
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
                throw new System.ArgumentException("formulas contains a null value");
            }


            if (formulas.Count != srcs.Count)
            {
                string msg = string.Format("Must have the same number of srcs ({0}) and formulas ({1})", srcs.Count,formulas.Count);
                throw new System.ArgumentException(msg);
            }


            var update = new VA.ShapeSheet.Update();
            update.BlastGuards  = ((short) flags & (short) IVisio.VisGetSetArgs.visSetBlastGuards)!=0;
            update.TestCircular = ((short) flags & (short) IVisio.VisGetSetArgs.visSetTestCircular) != 0;
            var shapeids = shapes.Select(s=>s.ID).ToList();

            int num_formulas = formulas.Count;
            foreach (var shapeid in shapeids)
            {
                for (int i=0; i<num_formulas;i++)
                {
                    var src = srcs[i];
                    var formula = formulas[i];
                    update.SetFormula((short) shapeid, src, formula);        
                }

            }

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Set ShapeSheet Formulas"))
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        public void SetResult(IList<IVisio.Shape> target_shapes, IList<VA.ShapeSheet.SRC> srcs,
               IList<double> results, IVisio.VisGetSetArgs flags)
        {
            var shapes = this.get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
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
                throw new System.ArgumentException("formulas contains a null value");
            }


            if (results.Count != srcs.Count)
            {
                string msg = string.Format("Must have the same number of srcs ({0}) and formulas ({1})", srcs.Count, results.Count);
                throw new System.ArgumentException(msg);
            }

            var update = new VA.ShapeSheet.Update();
            update.BlastGuards = ((short)flags & (short)IVisio.VisGetSetArgs.visSetBlastGuards) != 0;
            update.TestCircular = ((short)flags & (short)IVisio.VisGetSetArgs.visSetTestCircular) != 0;
            var shapeids = shapes.Select(s => s.ID).ToList();

            int num_formulas = results.Count;
            foreach (var shapeid in shapeids)
            {
                for (int i = 0; i < num_formulas; i++)
                {
                    var src = srcs[i];
                    var result = results[i];
                    update.SetResult((short)shapeid, src, result, IVisio.VisUnitCodes.visNoCast );
                }
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, "Set ShapeSheet Result"))
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        public void SetResult(IList<IVisio.Shape> target_shapes, IList<VA.ShapeSheet.SRC> srcs,
               IList<string> results, IVisio.VisGetSetArgs flags)
        {
            var shapes = this.get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
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
                throw new System.ArgumentException("formulas contains a null value");
            }


            if (results.Count != srcs.Count)
            {
                string msg = string.Format("Must have the same number of srcs ({0}) and formulas ({1})", srcs.Count, results.Count);
                throw new System.ArgumentException(msg);
            }

            var update = new VA.ShapeSheet.Update();
            update.BlastGuards = ((short)flags & (short)IVisio.VisGetSetArgs.visSetBlastGuards) != 0;
            update.TestCircular = ((short)flags & (short)IVisio.VisGetSetArgs.visSetTestCircular) != 0;
            var shapeids = shapes.Select(s => s.ID).ToList();

            int num_formulas = results.Count;
            foreach (var shapeid in shapeids)
            {
                for (int i = 0; i < num_formulas; i++)
                {
                    var src = srcs[i];
                    var result = results[i];
                    update.SetResult((short)shapeid, src, result, IVisio.VisUnitCodes.visNoCast);
                }
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, "Set ShapeSheet Result"))
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }


        public void Update(ShapeSheetUpdate update, bool blastguards, bool testcircular)
        {
            this.Session.WriteVerbose( "Staring ShapeSheet Update");
            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Update ShapeSheet Formulas"))
            {
                var active_page = application.ActivePage;
                var internal_update = update.update;
                internal_update.BlastGuards = blastguards;
                internal_update.TestCircular = testcircular;
                this.Session.WriteVerbose( "BlastGuards={0}", blastguards);
                this.Session.WriteVerbose( "TestCircular={0}", testcircular);
                internal_update.Execute(active_page);                
            }
            this.Session.WriteVerbose( "Ending ShapeSheet Update");
        }
    }
}