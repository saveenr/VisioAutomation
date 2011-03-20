using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Update
{
    public class FormulaData<TStream> where TStream : struct
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

        public IList<FormulaItem<TStream>> Items
        {
            get { return this.items; }
        }
    }
}