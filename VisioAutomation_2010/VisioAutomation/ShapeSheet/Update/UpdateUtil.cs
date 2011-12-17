using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Update
{
    internal static class UpdateUtil
    {
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
                flags = (IVisio.VisGetSetArgs) ((short) flags | (short) IVisio.VisGetSetArgs.visSetUniversalSyntax);
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
            short flags,
            int numitems)
        {
            if (numitems < 1)
            {
                return 0;
            }

            var formula_obj_array = UpdateUtil.StringsToObjectArray(formulas);

            // Force UniversalSyntax 
            flags |= (short) IVisio.VisGetSetArgs.visSetUniversalSyntax;

            return page.SetFormulas(stream, formula_obj_array, flags);
        }


        public static short SetFormulas(
            IVisio.Shape shape,
            short[] stream,
            IList<string> formulas,
            IVisio.VisGetSetArgs flags,
            int numitems)
        {
            if (formulas.Count != numitems)
            {
                string msg = string.Format("Expected {0} formulas, instead have {1}", numitems, formulas.Count);
                throw new AutomationException(msg);
            }

            if (numitems == 0)
            {
                return 0;
            }


            var formula_obj_array = UpdateUtil.StringsToObjectArray(formulas);

            // Force UniversalSyntax 
            short short_flags = (short) (((short) flags) | ((short) IVisio.VisGetSetArgs.visSetUniversalSyntax));

            return shape.SetFormulas(stream, formula_obj_array, short_flags);
        }

        public static short SetResults(
            IVisio.Shape shape,
            short[] stream,
            IList<double> results,
            IList<IVisio.VisUnitCodes> unit_codes,
            IVisio.VisGetSetArgs flags,
            int numitems)
        {
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
            var results_obj_array = UpdateUtil.DoublesToObjectArray(results);

            flags = UpdateUtil.CheckSetResultsFlags(flags);

            short num_set = shape.SetResults(stream, unitcodes_obj_array, results_obj_array, (short) flags);

            return num_set;
        }

        public static short SetResults(
            IVisio.Page page,
            short[] stream,
            IList<double> results,
            IList<IVisio.VisUnitCodes> unitcodes,
            IVisio.VisGetSetArgs flags,
            int numitems)
        {
            if (results.Count != numitems)
            {
                string msg = string.Format("Expected {0} results, instead have {1}", numitems, results.Count);
                throw new AutomationException(msg);
            }

            if (numitems == 0)
            {
                return 0;
            }

            var results_obj_array = UpdateUtil.DoublesToObjectArray(results);
            var unitcodes_obj_array = VA.ShapeSheet.ShapeSheetHelper.UnitCodesToObjectArray(unitcodes);

            flags = UpdateUtil.CheckSetResultsFlags(flags);

            return page.SetResults(stream, unitcodes_obj_array, results_obj_array, (short) flags);
        }
    }
}