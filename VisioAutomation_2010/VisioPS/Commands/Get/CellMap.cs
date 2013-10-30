using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioPS.Commands
{
    public class CellMap
    {
        Dictionary<string, SRC> dic;
        
        public CellMap()
        {
            this.dic = new Dictionary<string, SRC>(System.StringComparer.OrdinalIgnoreCase);
        }

        public VisioAutomation.ShapeSheet.SRC this[string name]
        {
            get { return this.dic[name]; }
            set { this.dic[name] = value; }
        }

        public Dictionary<string, SRC>.KeyCollection CellNames
        {
            get
            {
                return this.dic.Keys;
            }
        }

        public IEnumerable<string> ResolveName(string cellname)
        {
            if (cellname.Contains("*") || cellname.Contains("?"))
            {
                string pat = "^" + System.Text.RegularExpressions.Regex.Escape(cellname)
                    .Replace(@"\*", ".*").
                    Replace(@"\?", ".") + "$";

                var regex = new System.Text.RegularExpressions.Regex(pat, System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                foreach (string k in this.CellNames)
                {
                    if (regex.IsMatch(k))
                    {
                        yield return k;
                    }
                }
            }
            else
            {
                if (!this.dic.ContainsKey(cellname))
                {
                    throw new System.ArgumentException("cellname not defined in map");
                }
                yield return cellname;
            }
        }


    }
}