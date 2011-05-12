using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VAM = VisioAutomationMin;

namespace VisioAutomationMin
{
    public class SRCUpdate
    {
        private struct Item
        {
            public SRC src;
            public  string formula;

            public Item(SRC src, string formula)
            {
                this.src = src;
                this.formula = formula;
            }
        }

        private List<Item> items;


        public SRCUpdate()
        {
            this.items = new List<Item>();
        }

        public void SetFormula(SRC src, FormulaLiteral formula)
        {
            var item = new Item(src, formula.Value);
            this.items.Add(item);
        }

        public void SetFormula(short s, short r, short c, FormulaLiteral formula)
        {
            var src = new SRC(s, r, c);
            var item = new Item(src, formula.Value);
            this.items.Add(item);
        }

        public int Execute(IVisio.Shape shape, IVisio.VisGetSetArgs flags)
        {
            short xflags = (short)(flags | IVisio.VisGetSetArgs.visSetUniversalSyntax);
            short[] SID_SRCStream = new short[items.Count * 3];
            object[] formulas_objects = new object[items.Count];
            for (int i = 0; i < items.Count; i++)
            {
                SID_SRCStream[i * 3 + 0] = items[i].src.Section;
                SID_SRCStream[i * 3 + 0] = items[i].src.Row;
                SID_SRCStream[i * 3 + 0] = items[i].src.Cell;
                formulas_objects[i] = items[i].formula;
            }

            int count = shape.SetFormulas(SID_SRCStream, formulas_objects, xflags);
            return count;
        }
    }
}