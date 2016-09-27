using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Scripting.ShapeSheet
{
    public class CellValueDictionary : CellDictionary<string>
    {
        private readonly CellSRCDictionary srcmap;

        public CellValueDictionary(CellSRCDictionary srcmap, Dictionary<string,object> dictionary)
        {
            if (srcmap == null)
            {
                throw new System.ArgumentNullException(nameof(srcmap));
            }

            this.srcmap = srcmap;

            this.UpdateFrom(dictionary);
        }


        public SRC GetSRC(string name)
        {
            return this.srcmap[name];
        }

        public void UpdateFrom(Dictionary<string,object> from_dic)
        {
            if (from_dic == null)
            {
                throw new System.ArgumentNullException(nameof(from_dic));
            }

            // We are certain all the keys are strings
            foreach (var pair in from_dic)
            {
                string cellname = pair.Key;
                object cell_value_object = pair.Value;
                var cell_value_string = CellValueDictionary.value_to_string(cell_value_object, cellname);
                this.UpdateFrom(cellname,cell_value_string);
            }
        }

        public void UpdateFrom(string cellname,string cellvalue)
        {
            if (!this.srcmap.ContainsCell(cellname))
            {
                string message = string.Format("Cell \"{0}\" is not supported", cellname);
                throw new System.ArgumentOutOfRangeException(message);
            }

            if (cellvalue == null)
            {
                string message = string.Format("Cell {0} has a null value. Use a non-null value", cellname);
                throw new System.ArgumentOutOfRangeException(message);
            }

            this[cellname] = cellvalue;
        }


        private static string value_to_string(object value_o, string cellname)
        {
            var invariant_culture = CultureInfo.InvariantCulture;

            string result;
            if (value_o is string)
            {
                result = (string) value_o;
            }
            else if (value_o is int)
            {
                int value_int = (int) value_o;
                result = value_int.ToString(invariant_culture);
            }
            else if (value_o is float)
            {
                float value_float = (float) value_o;
                result = value_float.ToString(invariant_culture);
            }
            else if (value_o is double)
            {
                double value_double = (double) value_o;
                result = value_double.ToString(invariant_culture);
            }
            else
            {
                var value_type_name = value_o.GetType().FullName;
                string message = string.Format(invariant_culture, "Cell {0} has an unsupported type {1} ", cellname, value_type_name);
                throw new System.ArgumentOutOfRangeException(message);
            }
            return result;
        }
    }
}