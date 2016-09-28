using System.Collections.Generic;
using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Utilities
{
    public struct QueryHelpers
    {
        private static int check_stream_size(short[] stream, int chunksize)
        {
            if ((chunksize != 3) && (chunksize != 4))
            {
                throw new System.ArgumentOutOfRangeException("Chunksize must be 3 or 4");
            }

            int remainder = stream.Length % chunksize;

            if (remainder != 0)
            {
                string msg = string.Format("stream must have a multiple of {0} elements", chunksize);
                throw new System.ArgumentException(msg);
            }

            return stream.Length / chunksize;
        }

        public static string[] GetFormulasU_SIDSRC(ShapeSheetSurface surface, short[] stream)
        {
            int numitems = check_stream_size(stream, 4);
            var formulas = _GetFormulasU(surface, stream, numitems);
            return formulas;
        }

        public static string[] GetFormulasU_SRC(ShapeSheetSurface surface, short[] stream)
        {
            int numitems = check_stream_size(stream, 3);
            var formulas = _GetFormulasU(surface, stream, numitems);
            return formulas;
        }

        private static string[] _GetFormulasU(ShapeSheetSurface surface, short[] stream, int numitems)
        {
            if (numitems < 1)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;

            if (surface.Target.Master != null)
            {
                surface.Target.Master.GetFormulasU(stream, out formulas_sa);
            }
            else if (surface.Target.Page != null)
            {
                surface.Target.Page.GetFormulasU(stream, out formulas_sa);
            }
            else if (surface.Target.Shape != null)
            {
                surface.Target.Shape.GetFormulasU(stream, out formulas_sa);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            object[] formulas_obj_array = (object[])formulas_sa;

            if (formulas_obj_array.Length != numitems)
            {
                string msg = string.Format("Expected {0} items from GetFormulas but only received {1}", numitems,
                    formulas_obj_array.Length);
                throw new InternalAssertionException(msg);
            }

            string[] formulas = new string[formulas_obj_array.Length];
            formulas_obj_array.CopyTo(formulas, 0);
            return formulas;
        }

        public static TResult[] GetResults_SIDSRC<TResult>(ShapeSheetSurface surface, short[] stream, IList<IVisio.VisUnitCodes> unitcodes)
        {
            int numitems = check_stream_size(stream, 4);
            var results = _GetResults<TResult>(surface, stream, unitcodes, numitems);
            return results;
        }

        public static TResult[] GetResults_SRC<TResult>(ShapeSheetSurface surface, short[] stream, IList<IVisio.VisUnitCodes> unitcodes)
        {
            int numitems = check_stream_size(stream, 3);
            var results = _GetResults<TResult>(surface, stream, unitcodes, numitems);
            return results;
        }


        public static TResult[] _GetResults<TResult>(ShapeSheetSurface surface, short[] stream, IList<IVisio.VisUnitCodes> unitcodes, int numitems)
        {
            EnforceValidResultType(typeof(TResult));

            if (numitems < 1)
            {
                return new TResult[0];
            }

            var result_type = typeof(TResult);
            var unitcodes_obj_array = get_unit_code_obj_array(unitcodes);
            var flags = get_VisGetSetArgs(result_type);

            System.Array results_sa = null;

            if (surface.Target.Master != null)
            {
                surface.Target.Master.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            }
            else if (surface.Target.Page != null)
            {
                surface.Target.Page.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            }
            else if (surface.Target.Shape != null)
            {
                surface.Target.Shape.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Target");
            }

            if (results_sa.Length != numitems)
            {
                string msg = string.Format("Expected {0} items from GetResults but only received {1}", numitems,
                    results_sa.Length);
                throw new InternalAssertionException(msg);
            }

            var results = new TResult[results_sa.Length];
            results_sa.CopyTo(results, 0);
            return results;
        }

        private static IVisio.VisGetSetArgs get_VisGetSetArgs(System.Type type)
        {
            IVisio.VisGetSetArgs flags;
            if (type == typeof(int))
            {
                flags = IVisio.VisGetSetArgs.visGetTruncatedInts;
            }
            else if (type == typeof(double))
            {
                flags = IVisio.VisGetSetArgs.visGetFloats;
            }
            else if (type == typeof(string))
            {
                flags = IVisio.VisGetSetArgs.visGetStrings;
            }
            else
            {
                string msg = string.Format("Unsupported Result Type: {0}", type.Name);
                throw new InternalAssertionException(msg);
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

        private static void EnforceValidResultType(System.Type result_type)
        {
            if (!IsValidResultType(result_type))
            {
                string msg = string.Format("Unsupported Result Type: {0}", result_type.Name);
                throw new InternalAssertionException(msg);
            }
        }

        private static bool IsValidResultType(System.Type result_type)
        {
            return (result_type == typeof(int)
                    || result_type == typeof(double)
                    || result_type == typeof(string));
        }

    }
}