using System.Collections.Generic;
using System.Collections;

namespace VisioAutomation.ShapeSheet.Update
{
    public class FormulaData<TStream> : IEnumerable<FormulaItem<TStream>> where TStream : struct 
    {
        private readonly List<FormulaItem<TStream>> items;

        public FormulaData()
        {
            this.items = new List<FormulaItem<TStream>>();
        }

        public FormulaData(int capacity)
        {
            this.items = new List<FormulaItem<TStream>>(capacity);
        }

        public int Count
        {
            get { return this.items.Count; }
        }

        public void Set(TStream streamitem, FormulaLiteral literal)
        {
            ShapeSheetHelper.CheckFormulaIsNotNull(literal.Value);
            var rec = new FormulaItem<TStream>(streamitem, literal.Value);
            this.items.Add(rec);
        }

        public string[] GetFormulasArray()
        {
            return ShapeSheetHelper.MapCollectionToArray(this.items, r => r.Formula);
        }

        public IEnumerator<FormulaItem<TStream>> GetEnumerator()
        {
            foreach (var i in this.items)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     // Explicit implementation
        {                                           // keeps it hidden.
            return GetEnumerator();
        }
    }
}