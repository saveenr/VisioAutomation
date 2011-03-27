using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ShapeSheetCommands : SessionCommands
    {
        public ShapeSheetCommands(Session session) :
            base(session)
        {

        }

        public VA.ShapeSheet.Query.Table<T> QueryResults<T>(VA.ShapeSheet.SRC src)
        {
            var srcs = new[] { src };
            return QueryResults<T>(srcs);
        }

        public VA.ShapeSheet.Query.Table<T> QueryResults<T>(IList<VA.ShapeSheet.SRC> srcs)
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

        public VA.ShapeSheet.Query.Table<string> QueryFormulas(VA.ShapeSheet.SRC src)
        {
            var srcs = new[] { src };
            return QueryFormulas(srcs);
        }

        public VA.ShapeSheet.Query.Table<string> QueryFormulas(IList<VA.ShapeSheet.SRC> srcs)
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

        public VA.ShapeSheet.Query.Table<T> QueryResults<T>(IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
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

        public VA.ShapeSheet.Query.Table<string> QueryFormulas(IVisio.VisSectionIndices section, IList<IVisio.VisCellIndices> cells)
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

        /// <summary>
        /// Optimizes setting formulas for cells identified by names
        /// </summary>
        /// <param name="cellname"></param>
        /// <param name="formula"></param>
        /// <param name="flags"></param>
        public void SetFormula(string cellname, string formula, IVisio.VisGetSetArgs flags)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            VA.ShapeSheet.SRC? src = VA.ShapeSheet.ShapeSheetHelper.TryGetSRCFromName(cellname);
            if (src.HasValue)
            {
                // if cellrcs is one we have optimized for, we'll have its SRC vaue
                var srcs = new [] {src.Value};
                var formulas = new [] {formula };

                // simply call SetFormulas using the SRC value and everything will work fast
                SetFormula(srcs, formulas, flags);
            }
            else
            {
                // In this case, we didn't find a SRC value for the name
                // So we resort to setting the formulas, one-by-one
                // This is very slow, but it should not occur in practice very often
                var shapes = this.Session.Selection.GetSelectedShapes(ShapesEnumeration.Flat);
                foreach (var shape in shapes)
                {
                    var cell = shape.Cells[cellname];
                    cell.FormulaU = formula;
                }
            }
        }
        
        public void SetFormula(IList<VA.ShapeSheet.SRC> srcs, 
            IList<string> formulas,
            IVisio.VisGetSetArgs flags)
        {
            if (srcs == null)
            {
                throw new ArgumentNullException("srcs");
            }

            if (formulas == null)
            {
                throw new ArgumentNullException("formulas");
            }

            if (formulas.Any( f => f == null))
            {
                throw new ArgumentException("formulas contains a null value");
            }


            if (formulas.Count != srcs.Count)
            {
                string msg = string.Format("Must have the same number of srcs ({0}) and formulas ({1})", srcs.Count,formulas.Count);
                throw new ArgumentException(msg);
            }


            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            update.BlastGuards  = ((short) flags & (short) IVisio.VisGetSetArgs.visSetBlastGuards)!=0;
            update.TestCircular = ((short) flags & (short) IVisio.VisGetSetArgs.visSetTestCircular) != 0;
            var selection = this.Session.Selection.GetSelection();
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
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        public void SetCells(CellSetter cellsetter, bool blastguards, bool testcircular)
        {
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                var update = cellsetter.update;
                update.BlastGuards = blastguards;
                update.TestCircular = testcircular;

                update.Execute(active_page);                
            }
        }

        public void SetFormulas<T>(IEnumerable<T> items, 
            Func<T, bool> has_data,
            Func<T, VA.ShapeSheet.SRC> get_src,
            Func<T, string> get_formula)
        {
            var selection = this.Session.Selection.GetSelection();
            var shapeids = selection.GetIDs();
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            foreach (var shapeid in shapeids)
            {
                foreach (var item in items)
                {
                    if (has_data(item))
                    {
                        var src = get_src(item);
                        var formula = get_formula(item);
                        update.SetFormula((short)shapeid, src, formula);
                    }
                }
            }

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                update.Execute(active_page);
            }
        }

        public VA.ShapeSheet.Query.Table<string> QueryFormulas(IList<string> cellnames)
        {
            if (!this.Session.HasSelectedShapes())
            {
                throw new AutomationException("Needs at least 1 selected shape");
            }

            var srcs = this._CellNamesToSRCs(cellnames);
            var formulas = this.QueryFormulas(srcs);
            return formulas;
        }

        public VA.ShapeSheet.Query.Table<T> QueryResults<T>(IList<string> cellnames)
        {
            if (!this.Session.HasSelectedShapes())
            {
                throw new AutomationException("Needs at least 1 selected shape");
            }

            var srcs = this._CellNamesToSRCs(cellnames);
            var results = this.QueryResults<T>(srcs);
            return results;
        }

        private IList<VA.ShapeSheet.SRC> _CellNamesToSRCs(IList<string> cellnames)
        {
            var srcs = cellnames.Select(s => VA.ShapeSheet.ShapeSheetHelper.GetSRCFromName(s)).ToList();
            return srcs;
        }
    }
}