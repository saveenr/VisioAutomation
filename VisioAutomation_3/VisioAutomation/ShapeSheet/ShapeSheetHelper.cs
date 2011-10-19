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

            return MapCollectionToArray(unitcodes, uc => (object)uc);
        }




        internal static void CheckValidDataTypeForResult(Type result_type)
        {
            if (!((result_type == typeof(string)) || (result_type == typeof(int) || (result_type == typeof(double)))))
            {
                string msg = "type must be int, string or double";
                throw new ArgumentException(msg);
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
 


    }
}