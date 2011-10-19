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
    }
}