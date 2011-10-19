using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.Query
{
    internal static class QueryUtil
    {
        internal static IList<IVisio.VisUnitCodes> get_unitcodes_for_rows(IList<IVisio.VisUnitCodes> unitcodes, int rows)
        {
            var all_unitcodes = new List<IVisio.VisUnitCodes>(rows*unitcodes.Count);
            for (short row = 0; row < rows; row++)
            {
                all_unitcodes.AddRange(unitcodes);
            }
            return all_unitcodes;
        }

        internal static IVisio.VisGetSetArgs _GetResultsFlagForResultType(Type result_type)
        {
            IVisio.VisGetSetArgs flags = 0;

            if (result_type == typeof (int))
            {
                flags = IVisio.VisGetSetArgs.visGetTruncatedInts;
            }
            else if (result_type == typeof (double))
            {
                flags = IVisio.VisGetSetArgs.visGetFloats;
            }
            else if (result_type == typeof (string))
            {
                flags = IVisio.VisGetSetArgs.visGetStrings;
            }
            else
            {
                throw new ArgumentOutOfRangeException();
            }

            return flags;
        }

        internal static string[] GetFormulasU(
            IVisio.Page page,
            IList<SIDSRC> stream)
        {
            var array = VA.ShapeSheet.SIDSRC.ToStream(stream);
            return GetFormulasU(page, array, stream.Count);
        }


        internal static string[] GetFormulasU(
            IVisio.Page page,
            short[] stream, int numitems)
        {
            if (numitems == 0)
            {
                return new string[0];
            }

            Array formulas_sa;

            page.GetFormulasU(
                stream,
                out formulas_sa);

            object[] formulas_obj_array = (object[]) formulas_sa;

            if (formulas_obj_array.Length != numitems)
            {
                string msg = String.Format(
                    "Expected {0} items from GetFormulas but only received {1}",
                    numitems,
                    formulas_obj_array.Length);
                throw new AutomationException(msg);
            }

            string[] formulas = new string[formulas_obj_array.Length];
            formulas_obj_array.CopyTo(formulas, 0);

            return formulas;
        }

        internal static string[] GetFormulasU(
            IVisio.Shape shape,
            short[] stream, int numitems)
        {
            if (numitems < 1)
            {
                return new string[0];
            }

            Array formulas_sa;
            shape.GetFormulasU(stream, out formulas_sa);

            object[] formulas_obj_array = (object[]) formulas_sa;

            if (formulas_obj_array.Length != numitems)
            {
                string msg = String.Format(
                    "Expected {0} items from GetFormulas but only received {1}",
                    numitems,
                    formulas_obj_array.Length);
                throw new AutomationException(msg);
            }

            string[] formulas = new string[formulas_obj_array.Length];
            formulas_obj_array.CopyTo(formulas, 0);

            return formulas;
        }


        internal static TResult[] GetResults<TResult>(
            IVisio.Page page,
            short[] stream,
            IList<IVisio.VisUnitCodes> unitcodes,
            int numitems)
        {
            if (numitems == 0)
            {
                return new TResult[0];
            }

            var result_type = typeof (TResult);
            var flags = VA.ShapeSheet.Query.QueryUtil._GetResultsFlagForResultType(result_type);
            var unitcodes_obj_array = VA.ShapeSheet.ShapeSheetHelper.UnitCodesToObjectArray(unitcodes);

            Array results_sa;

            page.GetResults(
                stream,
                (short) flags,
                unitcodes_obj_array,
                out results_sa);

            object[] results_obj_array = (object[]) results_sa;

            if (results_obj_array.Length != numitems)
            {
                string msg = String.Format(
                    "Expected {0} items from GetResults but only received {1}",
                    numitems,
                    results_obj_array.Length);
                throw new AutomationException(msg);
            }

            TResult[] results = new TResult[results_obj_array.Length];
            results_obj_array.CopyTo(results, 0);

            return results;
        }


        internal static TResult[] GetResults<TResult>(
            IVisio.Shape shape,
            List<VA.ShapeSheet.SRC> stream,
            IList<IVisio.VisUnitCodes> unitcodes)
        {
            var array = VA.ShapeSheet.SRC.ToStream(stream);
            return GetResults<TResult>(shape, array, unitcodes, stream.Count);
        }

        internal static TResult[] GetResults<TResult>(
            IVisio.Shape shape,
            short[] stream,
            IList<IVisio.VisUnitCodes> unitcodes, int numitems)
        {
            if (numitems < 1)
            {
                return new TResult[0];
            }

            var result_type = typeof (TResult);
            var unitcodes_obj_array = VA.ShapeSheet.ShapeSheetHelper.UnitCodesToObjectArray(unitcodes);
            var flags = VA.ShapeSheet.Query.QueryUtil._GetResultsFlagForResultType(result_type);

            Array results_sa;
            shape.GetResults(
                stream,
                (short) flags,
                unitcodes_obj_array,
                out results_sa);

            object[] results_obj_array = (object[]) results_sa;

            if (results_obj_array.Length != numitems)
            {
                string msg = String.Format(
                    "Expected {0} items from GetResults but only received {1}",
                    numitems,
                    results_obj_array.Length);
                throw new AutomationException(msg);
            }

            TResult[] results = new TResult[results_obj_array.Length];
            results_obj_array.CopyTo(results, 0);

            return results;
        }
    }

    public class QueryBase<TCol> where TCol : QueryColumn
    {
        private QueryColumnList<TCol> _columns;

        internal QueryBase()
        {
            this._columns = new QueryColumnList<TCol>();
        }

        public QueryColumnList<TCol> Columns
        {
            get { return this._columns; }
        }

        protected void AddColumn(TCol column)
        {
            if (column == null)
            {
                throw new ArgumentNullException("column");
            }

            this._columns.Add(column);
        }

        protected IList<IVisio.VisUnitCodes> CreateUnitCodeArray()
        {
            var a = new IVisio.VisUnitCodes[this.Columns.Count];
            for (int i = 0; i < this.Columns.Count; i++)
            {
                a[i] = this.Columns[i].UnitCode;
            }

            return a;
        }

        protected void validate_unitcodes(IList<IVisio.VisUnitCodes> unitcodes, int total_cells)
        {
            if (unitcodes == null)
            {
                throw new AutomationException("unitcodes must not be null");
            }

            if (unitcodes.Count != total_cells)
            {
                string msg = string.Format("Expected {0} unitcodes", total_cells);
                throw new AutomationException(msg);
            }
        }
    }
}