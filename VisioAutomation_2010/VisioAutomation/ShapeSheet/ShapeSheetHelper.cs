using System;
using System.Collections.Generic;
using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet
{
    public static class ShapeSheetHelper
    {
        private static int check_stream_size(short[] stream, int chunksize)
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
            if (numitems < 1)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;
            page.GetFormulasU(stream, out formulas_sa);

            var formulas = get_formulas_array(formulas_sa, numitems);
            return formulas;
        }

        public static string[] GetFormulasU(IVisio.Shape shape, short[] stream)
        {
            int numitems = check_stream_size(stream, 3);
            if (numitems < 1)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;
            shape.GetFormulasU(stream, out formulas_sa);

            var formulas = get_formulas_array(formulas_sa, numitems);
            return formulas;
        }

        private static string[] get_formulas_array(Array formulas_sa, int numitems)
        {
            object[] formulas_obj_array = (object[]) formulas_sa;

            if (formulas_obj_array.Length != numitems)
            {
                string msg = string.Format(
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
            EnforceValidResultType(typeof(TResult));

            int numitems = check_stream_size(stream, 4);
            if (numitems < 1)
            {
                return new TResult[0];
            }

            var result_type = typeof(TResult);
            var unitcodes_obj_array = get_unit_code_obj_array(unitcodes);
            var flags = get_VisGetSetArgs(result_type);

            System.Array results_sa = null;
            page.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            var results = get_results_array<TResult>(results_sa, numitems);

            return results;
        }

        public static TResult[] GetResults<TResult>(IVisio.Shape shape, short[] stream, IList<IVisio.VisUnitCodes> unitcodes)
        {
            EnforceValidResultType(typeof(TResult));

            int numitems  = check_stream_size(stream, 3);

            if (numitems < 1)
            {
                return new TResult[0];
            }

            var result_type = typeof(TResult);
            var unitcodes_obj_array = get_unit_code_obj_array(unitcodes);
            var flags = get_VisGetSetArgs(result_type);

            System.Array results_sa = null;
            shape.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            var results = get_results_array<TResult>(results_sa, numitems);

            return results;
        }

        private static TResult[] get_results_array<TResult>(Array results_sa, int numitems)
        {
            if (results_sa.Length != numitems)
            {
                string msg = string.Format(
                    "Expected {0} items from GetResults but only received {1}",
                    numitems,
                    results_sa.Length);
                throw new AutomationException(msg);
            }

            TResult[] results = new TResult[results_sa.Length];
            results_sa.CopyTo(results, 0);
            return results;
        }

        private static IVisio.VisGetSetArgs get_VisGetSetArgs(Type type)
        {
            IVisio.VisGetSetArgs flags;
            if (type == typeof (int))
            {
                flags = IVisio.VisGetSetArgs.visGetTruncatedInts;
            }
            else if (type == typeof (double))
            {
                flags = IVisio.VisGetSetArgs.visGetFloats;
            }
            else if (type == typeof (string))
            {
                flags = IVisio.VisGetSetArgs.visGetStrings;
            }
            else
            {
                string msg = string.Format("Internal error: Unsupported Result Type: {0}", type.Name);
                throw new VA.AutomationException(msg);
            }
            return flags;
        }

        private static object[] get_unit_code_obj_array(IList<IVisio.VisUnitCodes> unitcodes)
        {
            // Create the unit codes array
            object[] unitcodes_obj_array = null;
            if (unitcodes != null)
            {
                unitcodes_obj_array = new object[unitcodes.Count];
                for (int i = 0; i < unitcodes.Count; i++)
                {
                    unitcodes_obj_array[i] = unitcodes[i];
                }
            }
            return unitcodes_obj_array;
        }

        internal static void EnforceValidResultType(System.Type result_type)
        {
            if (!IsValidResultType(result_type))
            {
                string msg = string.Format("Unsupported Result Type: {0}", result_type.Name);
                throw new VA.AutomationException(msg);
            }
        }

        internal static bool IsValidResultType(System.Type result_type)
        {
            return (result_type == typeof (int)
                    || result_type == typeof (double)
                    || result_type == typeof (string));
        }

    }
}