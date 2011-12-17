using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.Query
{
    static class QueryUtil
    {
        public static IVisio.VisGetSetArgs ResultTypeToGetResultsFlag(Type result_type)
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
        
        public static string[] GetFormulasU( IVisio.Page page, short[] stream, int numitems)
        {
            if (numitems == 0)
            {
                return new string[0];
            }

            Array formulas_sa;

            page.GetFormulasU(stream, out formulas_sa);

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

        public static string[] GetFormulasU( IVisio.Shape shape, short[] stream, int numitems)
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
        
        public static TResult[] GetResults<TResult>( IVisio.Page page, short[] stream, IList<IVisio.VisUnitCodes> unitcodes, int numitems)
        {
            if (numitems == 0)
            {
                return new TResult[0];
            }

            var result_type = typeof (TResult);
            var flags = VA.ShapeSheet.Query.QueryUtil.ResultTypeToGetResultsFlag(result_type);
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
        
        public static TResult[] GetResults<TResult>( IVisio.Shape shape, short[] stream, IList<IVisio.VisUnitCodes> unitcodes, int numitems)
        {
            if (numitems < 1)
            {
                return new TResult[0];
            }

            var result_type = typeof (TResult);
            var unitcodes_obj_array = VA.ShapeSheet.ShapeSheetHelper.UnitCodesToObjectArray(unitcodes);
            var flags = VA.ShapeSheet.Query.QueryUtil.ResultTypeToGetResultsFlag(result_type);

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
}