using System;
using System.Collections.Generic;
using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet
{
    public static partial class ShapeSheetHelper
    {
        private static object[] StringsToObjectArray(IList<string> strings)
        {
            if (strings == null)
            {
                return null;
            }

            return MapCollectionToArray(strings, uc => (object)uc);
        }

        private static object[] UnitCodesToObjectArray(IList<IVisio.VisUnitCodes> unitcodes)
        {
            if (unitcodes == null)
            {
                return null;
            }

            return MapCollectionToArray(unitcodes, uc => (object)uc);
        }

        private static object[] DoublesToObjectArray(IList<double> doubles)
        {
            if (doubles == null)
            {
                return null;
            }

            return MapCollectionToArray(doubles, uc => (object)uc);
        }

        private static IVisio.VisGetSetArgs _CheckSetResultsFlags(IVisio.VisGetSetArgs flags)
        {
            if ((flags & IVisio.VisGetSetArgs.visSetUniversalSyntax) > 0)
            {
                string msg = String.Format("visSetUniversalSyntax allowed only with visSetFormulas");
                throw new AutomationException(msg);
            }

            // force universal syntax if strings are set as formulas
            // if SetResults will fail if UniversalSyntax flag is used alone
            if ((flags & IVisio.VisGetSetArgs.visSetFormulas) > 0)
            {
                flags = (IVisio.VisGetSetArgs)((short)flags | (short)IVisio.VisGetSetArgs.visSetUniversalSyntax);
            }

            return flags;
        }

        internal static void CheckValidDataTypeForResult(Type result_type)
        {
            if (!((result_type == typeof(string)) || (result_type == typeof(int) || (result_type == typeof(double)))))
            {
                string msg = "type must be int, string or double";
                throw new ArgumentException(msg);
            }
        }

        internal static IVisio.VisGetSetArgs _GetResultsFlagForResultType(Type result_type)
        {
            IVisio.VisGetSetArgs flags = 0;

            if (result_type == typeof(int))
            {
                flags = IVisio.VisGetSetArgs.visGetTruncatedInts;
            }
            else if (result_type == typeof(double))
            {
                flags = IVisio.VisGetSetArgs.visGetFloats;
            }
            else if (result_type == typeof(string))
            {
                flags = IVisio.VisGetSetArgs.visGetStrings;
            }
            else
            {
                throw new ArgumentOutOfRangeException();
            }

            return flags;
        }

        internal static void CheckFormulaIsNotNull(string formula)
        {
            if (formula == null)
            {
                throw new AutomationException("Null not allowed for formula");
            }
        }

        internal static TB[] MapCollectionToArray<TA, TB>(IList<TA> source_collection, Func<TA, TB> xfrm)
        {
            int num_items = source_collection.Count;
            TB[] destination_array = new TB[num_items];
            for (int i = 0; i < num_items; i++)
            {
                destination_array[i] = xfrm(source_collection[i]);
            }

            return destination_array;
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
            short [] stream,
            IList<IVisio.VisUnitCodes> unitcodes, int numitems)
        {
            if (numitems < 1)
            {
                return new TResult[0];
            }

            var result_type = typeof(TResult);
            var unitcodes_obj_array = UnitCodesToObjectArray(unitcodes);
            var flags = _GetResultsFlagForResultType(result_type);

            Array results_sa;
            shape.GetResults(
                stream,
                (short)flags,
                unitcodes_obj_array,
                out results_sa);

            object[] results_obj_array = (object[])results_sa;

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

        internal static string[] GetFormulasU(
            IVisio.Shape shape,
            VA.ShapeSheet.Streams.SRCStream stream)
        {
            return GetFormulasU(shape, stream.Array, stream.Count);
        }

        internal static string[] GetFormulasU(
    IVisio.Shape shape,
    short [] stream, int numitems)
        {
            if (numitems< 1)
            {
                return new string[0];
            }

            Array formulas_sa;
            shape.GetFormulasU(stream, out formulas_sa);

            object[] formulas_obj_array = (object[])formulas_sa;

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

        internal static short SetFormulas(
    IVisio.Page page,
    VA.ShapeSheet.Streams.SIDSRCStream stream,
    IList<string> formulas,
    short flags)
        {
            return SetFormulas(page, stream.Array, formulas, flags, stream.Count);
        }


        internal static short SetFormulas(
            IVisio.Page page,
            short[] stream,
            IList<string> formulas,
            short flags,
            int numitems)
        {
            if (numitems< 1)
            {
                return 0;
            }

            var formula_obj_array = StringsToObjectArray(formulas);

            // Force UniversalSyntax 
            flags |= (short)IVisio.VisGetSetArgs.visSetUniversalSyntax;

            return page.SetFormulas(stream, formula_obj_array, flags);
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
            short [] stream, int numitems)
        {
            if (numitems == 0)
            {
                return new string[0];
            }

            Array formulas_sa;

            page.GetFormulasU(
                stream,
                out formulas_sa);

            object[] formulas_obj_array = (object[])formulas_sa;

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


        internal static short SetFormulas(
            IVisio.Shape shape,
            VA.ShapeSheet.Streams.SRCStream stream,
            IList<string> formulas,
            IVisio.VisGetSetArgs flags)
        {
            return SetFormulas(shape, stream.Array, formulas, flags, stream.Count);
        }

        internal static short SetFormulas(
    IVisio.Shape shape,
    short [] stream,
    IList<string> formulas,
    IVisio.VisGetSetArgs flags,
            int numitems)
        {
            if (formulas.Count != numitems)
            {
                string msg = String.Format("Expected {0} formulas, instead have {1}", numitems, formulas.Count);
                throw new AutomationException(msg);
            }

            if (numitems == 0)
            {
                return 0;
            }


            var formula_obj_array = StringsToObjectArray(formulas);

            // Force UniversalSyntax 
            short short_flags = (short)(((short)flags) | ((short)IVisio.VisGetSetArgs.visSetUniversalSyntax));

            return shape.SetFormulas(stream, formula_obj_array, short_flags);
        }


        internal static short SetResults(
            IVisio.Shape shape,
            VA.ShapeSheet.Streams.SRCStream stream,
            IList<double> results,
            IList<IVisio.VisUnitCodes> unit_codes,
            IVisio.VisGetSetArgs flags)
        {
            return SetResults(shape, stream.Array, results, unit_codes, flags, stream.Count);
        }

        internal static short SetResults(
    IVisio.Shape shape,
    short [] stream,
    IList<double> results,
    IList<IVisio.VisUnitCodes> unit_codes,
    IVisio.VisGetSetArgs flags,
            int numitems)
        {
            if (unit_codes.Count != numitems)
            {
                string msg = String.Format("Expected {0} unit_codes, instead have {1}", numitems, unit_codes.Count);
                throw new AutomationException(msg);
            }

            if (results.Count != numitems)
            {
                string msg = String.Format("Expected {0} results, instead have {1}", numitems, results.Count);
                throw new AutomationException(msg);
            }

            if (numitems< 1)
            {
                return 0;
            }

            var unitcodes_obj_array = UnitCodesToObjectArray(unit_codes);
            var results_obj_array = DoublesToObjectArray(results);

            flags = _CheckSetResultsFlags(flags);

            short num_set = shape.SetResults(stream, unitcodes_obj_array, results_obj_array, (short)flags);

            return num_set;
        }


        internal static short SetResults(
            IVisio.Page page,
            VA.ShapeSheet.Streams.SIDSRCStream stream,
            IList<double> results,
            IList<IVisio.VisUnitCodes> unitcodes,
            IVisio.VisGetSetArgs flags)
        {
            return SetResults(page, stream.Array, results, unitcodes, flags, stream.Count);
        }

        internal static short SetResults(
    IVisio.Page page,
    short[] stream,
    IList<double> results,
    IList<IVisio.VisUnitCodes> unitcodes,
    IVisio.VisGetSetArgs flags,
            int numitems)
        {
            if (results.Count != numitems)
            {
                string msg = String.Format("Expected {0} results, instead have {1}", numitems, results.Count);
                throw new AutomationException(msg);
            }

            if (numitems == 0)
            {
                return 0;
            }

            var results_obj_array = DoublesToObjectArray(results);
            var unitcodes_obj_array = UnitCodesToObjectArray(unitcodes);

            flags = _CheckSetResultsFlags(flags);

            return page.SetResults(stream , unitcodes_obj_array, results_obj_array, (short)flags);
        }


        internal static TResult[] GetResults<TResult>(
            IVisio.Page page,
            VA.ShapeSheet.Streams.SIDSRCStream stream,
            IList<IVisio.VisUnitCodes> unitcodes)
        {
            if (stream.Count == 0)
            {
                return new TResult[0];
            }

            var result_type = typeof(TResult);
            var flags = _GetResultsFlagForResultType(result_type);
            var unitcodes_obj_array = UnitCodesToObjectArray(unitcodes);

            Array results_sa;

            page.GetResults(
                stream.Array,
                (short)flags,
                unitcodes_obj_array,
                out results_sa);

            object[] results_obj_array = (object[])results_sa;

            if (results_obj_array.Length != stream.Count)
            {
                string msg = String.Format(
                    "Expected {0} items from GetResults but only received {1}",
                    stream.Count,
                    results_obj_array.Length);
                throw new AutomationException(msg);
            }

            TResult[] results = new TResult[results_obj_array.Length];
            results_obj_array.CopyTo(results, 0);

            return results;
        }

        internal static TResult[] GetResults<TResult>(
    IVisio.Page page,
    short[] stream,
    IList<IVisio.VisUnitCodes> unitcodes,
            int numitems)
        {
            if (numitems== 0)
            {
                return new TResult[0];
            }

            var result_type = typeof(TResult);
            var flags = _GetResultsFlagForResultType(result_type);
            var unitcodes_obj_array = UnitCodesToObjectArray(unitcodes);

            Array results_sa;

            page.GetResults(
                stream,
                (short)flags,
                unitcodes_obj_array,
                out results_sa);

            object[] results_obj_array = (object[])results_sa;

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


        internal static IList<IVisio.VisUnitCodes> get_unitcodes_for_rows(IList<IVisio.VisUnitCodes> unitcodes, int rows)
        {
            var all_unitcodes = new List<IVisio.VisUnitCodes>(rows*unitcodes.Count);
            for (short row = 0; row < rows; row++)
            {
                all_unitcodes.AddRange(unitcodes);
            }
            return all_unitcodes;
        }
    }
}