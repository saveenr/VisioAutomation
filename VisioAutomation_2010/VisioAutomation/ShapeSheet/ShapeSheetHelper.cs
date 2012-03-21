using System;
using System.Collections.Generic;
using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet
{
    public static partial class ShapeSheetHelper
    {
        internal static object[] UnitCodesToObjectArray(IList<IVisio.VisUnitCodes> unitcodes)
        {
            if (unitcodes == null)
            {
                return null;
            }

            int num_items = unitcodes.Count;
            var destination_array = new object[num_items];
            for (int i = 0; i < num_items; i++)
            {
                destination_array[i] = unitcodes[i];
            }
            return destination_array;
        }

        internal static void CheckValidDataTypeForResult(Type result_type)
        {
            if (!((result_type == typeof(string)) || (result_type == typeof(int) || (result_type == typeof(double)))))
            {
                string msg = "type must be int, string or double";
                throw new ArgumentException(msg);
            }
        }

        public static IVisio.VisGetSetArgs CheckSetResultsFlags(IVisio.VisGetSetArgs flags)
        {
            if ((flags & IVisio.VisGetSetArgs.visSetUniversalSyntax) > 0)
            {
                string msg = string.Format("visSetUniversalSyntax allowed only with visSetFormulas");
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

        public static object[] StringsToObjectArray(IList<string> strings)
        {
            if (strings == null)
            {
                return null;
            }

            int num_items = strings.Count;
            var destination_array = new object[num_items];
            for (int i = 0; i < num_items; i++)
            {
                destination_array[i] = strings[i];
            }
            return destination_array;
        }


        public static object[] DoublesToObjectArray(IList<double> doubles)
        {
            if (doubles == null)
            {
                return null;
            }

            int num_items = doubles.Count;
            var destination_array = new object[num_items];
            for (int i = 0; i < num_items; i++)
            {
                destination_array[i] = doubles[i];
            }
            return destination_array;
        }

        public static short SetFormulas(
            IVisio.Page page,
            short[] stream,
            IList<string> formulas,
            short flags)
        {
            int numitems = VA.ShapeSheet.ShapeSheetHelper.check_stream_size(stream, 4);

            if (numitems < 1)
            {
                return 0;
            }

            var formula_obj_array = VA.ShapeSheet.ShapeSheetHelper.StringsToObjectArray(formulas);

            // Force UniversalSyntax 
            flags |= (short)IVisio.VisGetSetArgs.visSetUniversalSyntax;

            return page.SetFormulas(stream, formula_obj_array, flags);
        }


        public static short SetFormulas(
            IVisio.Shape shape,
            short[] stream,
            IList<string> formulas,
            IVisio.VisGetSetArgs flags)
        {
            int numitems = VA.ShapeSheet.ShapeSheetHelper.check_stream_size(stream, 3);

            if (formulas.Count != numitems)
            {
                string msg = string.Format("Expected {0} formulas, instead have {1}", numitems, formulas.Count);
                throw new AutomationException(msg);
            }

            if (numitems == 0)
            {
                return 0;
            }


            var formula_obj_array = VA.ShapeSheet.ShapeSheetHelper.StringsToObjectArray(formulas);

            // Force UniversalSyntax 
            short short_flags = (short)(((short)flags) | ((short)IVisio.VisGetSetArgs.visSetUniversalSyntax));

            return shape.SetFormulas(stream, formula_obj_array, short_flags);
        }

        public static short SetResults(
            IVisio.Shape shape,
            short[] stream,
            IList<double> results,
            IList<IVisio.VisUnitCodes> unit_codes,
            IVisio.VisGetSetArgs flags)
        {
            int numitems = VA.ShapeSheet.ShapeSheetHelper.check_stream_size(stream, 3);

            if (unit_codes.Count != numitems)
            {
                string msg = string.Format("Expected {0} unit_codes, instead have {1}", numitems, unit_codes.Count);
                throw new AutomationException(msg);
            }

            if (results.Count != numitems)
            {
                string msg = string.Format("Expected {0} results, instead have {1}", numitems, results.Count);
                throw new AutomationException(msg);
            }

            if (numitems < 1)
            {
                return 0;
            }

            var unitcodes_obj_array = VA.ShapeSheet.ShapeSheetHelper.UnitCodesToObjectArray(unit_codes);
            var results_obj_array = VA.ShapeSheet.ShapeSheetHelper.DoublesToObjectArray(results);

            flags = VA.ShapeSheet.ShapeSheetHelper.CheckSetResultsFlags(flags);

            short num_set = shape.SetResults(stream, unitcodes_obj_array, results_obj_array, (short)flags);

            return num_set;
        }

        public static short SetResults(
            IVisio.Page page,
            short[] stream,
            IList<double> results,
            IList<IVisio.VisUnitCodes> unitcodes,
            IVisio.VisGetSetArgs flags)
        {
            int numitems = VA.ShapeSheet.ShapeSheetHelper.check_stream_size(stream, 4);

            if (results.Count != numitems)
            {
                string msg = string.Format("Expected {0} results, instead have {1}", numitems, results.Count);
                throw new AutomationException(msg);
            }

            if (numitems == 0)
            {
                return 0;
            }

            var results_obj_array = VA.ShapeSheet.ShapeSheetHelper.DoublesToObjectArray(results);
            var unitcodes_obj_array = VA.ShapeSheet.ShapeSheetHelper.UnitCodesToObjectArray(unitcodes);

            flags = VA.ShapeSheet.ShapeSheetHelper.CheckSetResultsFlags(flags);

            return page.SetResults(stream, unitcodes_obj_array, results_obj_array, (short)flags);
        }

        public static IVisio.VisGetSetArgs ResultTypeToGetResultsFlag(Type result_type)
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

        internal static int check_stream_size(short[] stream, int chunksize)
        {
            if ((chunksize != 3) && (chunksize != 4))
            {
                throw new VA.AutomationException("Chunksize must be 3 or 4");
            }

            int remainder = stream.Length % chunksize;

            if (remainder != 0)
            {
                string msg = string.Format("stream must have a multiple of {0} elements", chunksize);
                throw new VA.AutomationException(msg);
            }

            return stream.Length / chunksize;
        }

        public static string[] GetFormulasU(IVisio.Page page, short[] stream)
        {
            int numitems = check_stream_size(stream, 4);

            if (numitems == 0)
            {
                return new string[0];
            }

            Array formulas_sa;

            page.GetFormulasU(stream, out formulas_sa);

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

        public static string[] GetFormulasU(IVisio.Shape shape, short[] stream)
        {
            int numitems = check_stream_size(stream, 3);

            if (numitems < 1)
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

        public static TResult[] GetResults<TResult>(IVisio.Page page, short[] stream, IList<IVisio.VisUnitCodes> unitcodes)
        {
            int numitems = check_stream_size(stream, 4);

            if (numitems == 0)
            {
                return new TResult[0];
            }

            var result_type = typeof(TResult);
            var flags = VA.ShapeSheet.ShapeSheetHelper.ResultTypeToGetResultsFlag(result_type);
            var unitcodes_obj_array = VA.ShapeSheet.ShapeSheetHelper.UnitCodesToObjectArray(unitcodes);

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

        public static TResult[] GetResults<TResult>(IVisio.Shape shape, short[] stream, IList<IVisio.VisUnitCodes> unitcodes)
        {
            int numitems = check_stream_size(stream, 3);

            if (numitems < 1)
            {
                return new TResult[0];
            }

            var result_type = typeof(TResult);
            var unitcodes_obj_array = VA.ShapeSheet.ShapeSheetHelper.UnitCodesToObjectArray(unitcodes);
            var flags = VA.ShapeSheet.ShapeSheetHelper.ResultTypeToGetResultsFlag(result_type);

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

    }
}