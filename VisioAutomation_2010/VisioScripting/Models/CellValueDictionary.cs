using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioScripting.Models
{
    public class CellValueDictionary : NamedDictionary<string>
    {
        private readonly NamedSrcDictionary srcmap;

        public CellValueDictionary(NamedSrcDictionary srcmap, Dictionary<string,string> dictionary)
        {
            if (srcmap == null)
            {
                throw new System.ArgumentNullException(nameof(srcmap));
            }

            this.srcmap = srcmap;

            this.Update(dictionary);
        }


        public Src GetSrc(string name)
        {
            return this.srcmap[name];
        }

        public void Update(Dictionary<string,string> cells)
        {
            if (cells == null)
            {
                throw new System.ArgumentNullException(nameof(cells));
            }

            // We are certain all the keys are strings
            foreach (var pair in cells)
            {
                string cellname = pair.Key;
                this.Update(cellname, pair.Value);
            }
        }

        public void Update(string cellname, string cellvalue)
        {
            if (!this.srcmap.ContainsKey(cellname))
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
    }
}