using System.Collections;
using System.Globalization;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell
{
    public class CellValueDictionary : CellNameDictionary<string>
    {
        private readonly CellSRCDictionary srcmap;

        public CellValueDictionary(CellSRCDictionary srcmap) : base()
        {
            if (this.srcmap == null)
            {
                throw new System.ArgumentNullException("srcmap");
            }

            this.srcmap = srcmap;
        }

        public CellValueDictionary(CellSRCDictionary srcmap, Hashtable ht)
            : this(srcmap)
        {

            this.UpdateValueMap(ht);
        }


        public SRC GetSRC(string name)
        {
            return this.srcmap[name];
        }

        public void UpdateValueMap(Hashtable ht)
        {
            if (ht == null)
            {
                throw new System.ArgumentNullException("ht");
            }

            // Validate that all the keys are strings
            foreach (object key_o in ht.Keys)
            {
                if (!(key_o is string))
                {
                    string message = string.Format("Only string values can be keys in the hashtable. Encountered a key of type {0}", key_o.GetType().FullName);
                    throw new System.ArgumentOutOfRangeException("ht", message);
                }
            }


            // We are certain all the keys are strings
            foreach (object key_o in ht.Keys)
            {
                string cellname = (string) key_o;

                if (!srcmap.ContainsCell(cellname))
                {
                    string message = string.Format("Cell \"{0}\" is not supported", cellname);
                    throw new System.ArgumentOutOfRangeException("ht", message);                    
                }
                var cell_value_o = ht[key_o];

                if (cell_value_o == null)
                {
                    string message = string.Format("Cell {0} has a null value. Use a non-null value", cellname);
                    throw new System.ArgumentOutOfRangeException("ht", message);
                }

                var cell_value_string = CellValueDictionary.get_value_string(cell_value_o, cellname);

                this[cellname] = cell_value_string;
            }
        }

        private static string get_value_string(object value_o, string cellname)
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
                var value_type_name = value_o.GetType().FullName;
                string message = string.Format("Cell {0} has an unsupported type {1} ", cellname, value_type_name);
                throw new System.ArgumentOutOfRangeException(message);
            }
            return value_string;
        }
    }
}