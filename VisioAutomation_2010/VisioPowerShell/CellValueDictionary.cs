
using System;
using System.Collections;
using System.Globalization;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell
{
    public class CellValueDictionary : CellNameDictionary<string>
    {
        private readonly CellSRCDictionary srcmap;

        public CellValueDictionary( CellSRCDictionary src_src_dictionary) : base()
        {
            this.srcmap = src_src_dictionary;
        }

        public SRC GetSRC(string name)
        {
            return this.srcmap[name];
        }

        public void UpdateValueMap(Hashtable ht, CellSRCDictionary cellmap)
        {
            if (ht == null)
            {
                throw new System.ArgumentNullException("ht");
            }

            if (cellmap == null)
            {
                throw new System.ArgumentNullException("cellmap");
            }

            // Validate that all the keys are strings
            foreach (object key_o in ht.Keys)
            {
                if (!(key_o is string))
                {
                    string message = "Only string values can be keys in the hashtable";
                    throw new System.ArgumentOutOfRangeException("ht", message);
                }
            }

            // Validate that all the keys are strings
            foreach (object values_o in ht.Values)
            {
                if (values_o == null)
                {
                    string message = "Null values not allowed for cellvalues";
                    throw new System.ArgumentOutOfRangeException("ht", message);
                }
            }

            // We are certain all the keys are strings and all the values are not null
            foreach (object key_o in ht.Keys)
            {
                string key_string = (string) key_o;

                if (!cellmap.ContainsCell(key_string))
                {
                    string message = string.Format("Cell \"{0}\" is not supported", key_string);
                    throw new System.ArgumentOutOfRangeException("ht", message);                    
                }
                var value_o = ht[key_o];
                var value_string = get_value_string(value_o);
            }
        }

        private static string get_value_string(object value_o)
        {
            var culture = CultureInfo.InvariantCulture;

            string value_string;
            if (value_o is string)
            {
                value_string = (string) value_o;
            }
            else if (value_o is int)
            {
                int value_int = (int) value_o;
                value_string = value_int.ToString(culture);
            }
            else if (value_o is float)
            {
                float value_float = (float) value_o;
                value_string = value_float.ToString(culture);
            }
            else if (value_o is double)
            {
                double value_double = (double) value_o;
                value_string = value_double.ToString(culture);
            }
            else
            {
                string message = string.Format("Cell values cannot be of type {0} ", value_o.GetType().Name);
                throw new ArgumentOutOfRangeException(message);
            }
            return value_string;
        }
    }
}