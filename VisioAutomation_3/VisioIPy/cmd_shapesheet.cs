using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public VA.Scripting.CellSetter GetCellSetter()
        {
            return new VA.Scripting.CellSetter();
        }

        public void SetFormula(string cellname, string formula)
        {
            IVisio.VisGetSetArgs flags = 0;
            this.ScriptingSession.ShapeSheet.SetFormula(cellname, formula, flags);
        }

        public void SetFormulaSRC(VA.ShapeSheet.SRC src, string formula)
        {
            IVisio.VisGetSetArgs flags = 0;
            var formulas = new List<string> { formula };
            var srcs = new[] { src };
            this.ScriptingSession.ShapeSheet.SetFormula(srcs, formulas, flags);
        }

        public void SetCells(VA.Scripting.CellSetter setter)
        {
            bool blastguards = true;
            bool testcircular = false;
            this.ScriptingSession.ShapeSheet.SetCells(setter, blastguards, testcircular);
        }

        public VA.ShapeSheet.Query.Table<string> GetFormulasSRC(IList<VA.ShapeSheet.SRC> srcs)
        {
            var scriptingsession = this.ScriptingSession;
            var table = scriptingsession.ShapeSheet.QueryFormulas(srcs);
            return table;
        }

        public VA.ShapeSheet.Query.Table<string> GetFormulas(IList<string> cellnames)
        {
            var scriptingsession = this.ScriptingSession;
            var formulas = scriptingsession.ShapeSheet.QueryFormulas(cellnames);
            return formulas;
        }

        public VA.ShapeSheet.Query.Table<T> GetResultsSRC<T>(IList<VA.ShapeSheet.SRC> srcs)
        {
            var scriptingsession = this.ScriptingSession;
            var results = scriptingsession.ShapeSheet.QueryResults<T>(srcs);
            return results;
        }

        public VA.ShapeSheet.Query.Table<T> GetResults<T>(IList<string> cellnames)
        {
            var scriptingsession = this.ScriptingSession;
            var table = scriptingsession.ShapeSheet.QueryResults<T>(cellnames);
            return table;
        }

        public IList<VA.Layout.XFormCells> GetXFormData()
        {
            var scriptingsession = this.ScriptingSession;
            var data = scriptingsession.Layout.GetXForm();
            return data;
        }
    }
}