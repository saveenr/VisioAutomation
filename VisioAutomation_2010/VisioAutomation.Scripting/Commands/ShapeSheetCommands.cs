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

        public VA.ShapeSheet.Data.Table<T> QueryResults<T>(VA.ShapeSheet.SRC src)
        {
            var srcs = new[] { src };
            return QueryResults<T>(srcs);
        }

        public VA.ShapeSheet.Data.Table<T> QueryResults<T>(IList<VA.ShapeSheet.SRC> srcs)
        {
            var app = this.Session.VisioApplication;
            var page = app.ActivePage;
            var active_window = app.ActiveWindow;
            var selection = active_window.Selection;
            var shapeids = selection.GetIDs();

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

        public VA.ShapeSheet.Data.Table<string> QueryFormulas(VA.ShapeSheet.SRC src)
        {
            var srcs = new[] { src };
            return QueryFormulas(srcs);
        }

        public VA.ShapeSheet.Data.Table<string> QueryFormulas(IList<VA.ShapeSheet.SRC> srcs)
        {
            var app = this.Session.VisioApplication;
            var page = app.ActivePage;
            var active_window = app.ActiveWindow;
            var selection = active_window.Selection;
            var shapeids = selection.GetIDs();

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

        public VA.ShapeSheet.Data.Table<T> QueryResults<T>(IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            var app = this.Session.VisioApplication;
            var page = app.ActivePage;
            var active_window = app.ActiveWindow;
            var selection = active_window.Selection;
            var shapeids = selection.GetIDs();

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

        public VA.ShapeSheet.Data.Table<string> QueryFormulas(IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
        {
            var app = this.Session.VisioApplication;
            var page = app.ActivePage;
            var active_window = app.ActiveWindow;
            var selection = active_window.Selection;
            var shapeids = selection.GetIDs();

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
        
        public void SetFormula(IList<VA.ShapeSheet.SRC> srcs, 
            IList<string> formulas,
            IVisio.VisGetSetArgs flags)
        {
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


            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var update = new VA.ShapeSheet.Update();
            update.BlastGuards  = ((short) flags & (short) IVisio.VisGetSetArgs.visSetBlastGuards)!=0;
            update.TestCircular = ((short) flags & (short) IVisio.VisGetSetArgs.visSetTestCircular) != 0;
            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();

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