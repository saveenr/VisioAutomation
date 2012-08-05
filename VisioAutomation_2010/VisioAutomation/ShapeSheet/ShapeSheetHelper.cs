using System.Collections.Generic;
using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet
{
    public static partial class ShapeSheetHelper
    {
        private static object[] UnitCodesToObjectArray(IList<IVisio.VisUnitCodes> unitcodes)
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

        private static IVisio.VisGetSetArgs ResultTypeToGetResultsFlag(System.Type result_type)
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
                throw new System.ArgumentOutOfRangeException();
            }

            return flags;
        }

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
            return _GetFormulasU(page, stream);
        }

        public static string[] GetFormulasU(IVisio.Shape shape, short[] stream)
        {
            return _GetFormulasU(shape, stream);
        }

        public static string[] _GetFormulasU(object visio_object, short[] stream)
        {
            if (!(visio_object is IVisio.Page || visio_object is IVisio.Shape))
            {
                throw new VA.AutomationException("Internal error: Only Page and Shape objects supported in Execute()");
            }

            int numitems = -1; 

            if (visio_object is IVisio.Shape)
            {
                numitems = check_stream_size(stream, 3);
            }
            else if (visio_object is IVisio.Page)
            {
                numitems = check_stream_size(stream, 4);
            }

            if (numitems < 1)
            {
                return new string[0];
            }

            System.Array formulas_sa=null;

            if (visio_object is IVisio.Shape)
            {
                var shape = (IVisio.Shape)visio_object;
                shape.GetFormulasU(stream, out formulas_sa);
            }
            else if (visio_object is IVisio.Page)
            {
                var page = (IVisio.Page)visio_object;
                page.GetFormulasU(stream, out formulas_sa);
            }


            object[] formulas_obj_array = (object[])formulas_sa;

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
            return _GetResults<TResult>(page, stream, unitcodes);
        }

        public static TResult[] GetResults<TResult>(IVisio.Shape shape, short[] stream, IList<IVisio.VisUnitCodes> unitcodes)
        {
            return _GetResults<TResult>(shape, stream, unitcodes);
        }

        public static TResult[] _GetResults<TResult>(object visio_object, short[] stream, IList<IVisio.VisUnitCodes> unitcodes)
        {
            if (!(visio_object is IVisio.Page || visio_object is IVisio.Shape))
            {
                throw new VA.AutomationException("Internal error: Only Page and Shape objects supported in Execute()");
            }

            int numitems = -1; 

            if (visio_object is IVisio.Shape)
            {
                numitems = check_stream_size(stream, 3);
            }
            else if (visio_object is IVisio.Page)
            {
                numitems = check_stream_size(stream, 4);
            }
            
            if (numitems < 1)
            {
                return new TResult[0];
            }

            var result_type = typeof(TResult);
            var unitcodes_obj_array = VA.ShapeSheet.ShapeSheetHelper.UnitCodesToObjectArray(unitcodes);
            var flags = VA.ShapeSheet.ShapeSheetHelper.ResultTypeToGetResultsFlag(result_type);

            System.Array results_sa=null;
            if (visio_object is IVisio.Shape)
            {
                var shape = (IVisio.Shape) visio_object;
                shape.GetResults(stream, (short)flags, unitcodes_obj_array, out results_sa);
            }
            else if (visio_object is IVisio.Page)
            {
                var page = (IVisio.Page)visio_object;
                page.GetResults( stream, (short)flags, unitcodes_obj_array, out results_sa);
            }

            object[] results_obj_array = (object[])results_sa;

            if (results_obj_array.Length != numitems)
            {
                string msg = string.Format(
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