
namespace VisioPowerShell
{
    public class CellValueDictionary : CellNameDictionary<string>
    {
        private CellSRCDictionary srcmap;

        public CellValueDictionary( CellSRCDictionary src_src_dictionary) : base()
        {
            this.srcmap = src_src_dictionary;
        }

        public VisioAutomation.ShapeSheet.SRC GetSRC(string name)
        {
            return this.srcmap[name];
        }

        public void UpdateValueMap(System.Collections.Hashtable Hashtable)
        {
            if (Hashtable != null)
            {
                foreach (object key_o in Hashtable.Keys)
                {
                    if (!(key_o is string))
                    {
                        string message = "Only string values can be keys in the hashtable";
                        throw new System.ArgumentOutOfRangeException(message);
                    }
                    string key_string = (string)key_o;

                    object value_o = Hashtable[key_o];
                    if (value_o == null)
                    {
                        string message = "Null values not allowed for cellvalues";
                        throw new System.ArgumentOutOfRangeException(message);
                    }

                    var value_string = get_value_string(value_o);
                    this[key_string] = value_string;
                }
            }
        }

        private static string get_value_string(object value_o)
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;

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
                throw new System.ArgumentOutOfRangeException(message);
            }
            return value_string;
        }
    }
}