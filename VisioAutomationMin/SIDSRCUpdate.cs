using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VAM = VisioAutomationMin;

namespace VisioAutomationMin
{
    public class SIDSRCUpdate
    {
        private struct Item
        {
            public short id;
            public SRC src;
            public  string formula;

            public Item(short id, SRC src, string formula)
            {
                this.id = id;
                this.src = src;
                this.formula = formula;
            }
        }

        private List<Item> items;


        public SIDSRCUpdate()
        {
            this.items = new List<Item>();
        }

        public void SetFormula(short id, SRC src, FormulaLiteral formula)
        {
            var item = new Item(id, src, formula.Value);
            this.items.Add(item);
        }

        public void SetFormulaChecked(short id, SRC src, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                SetFormula(id,src,formula);
            }
        }

        public void SetFormula(short id, short s, short r, short c, FormulaLiteral formula)
        {
            var src = new SRC(s, r, c);
            var item = new Item(id, src, formula.Value);
            this.items.Add(item);
        }

        public void SetFormulaEx(short id, short s, short r, short c, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                SetFormula(id, s, r, c , formula);
            }
        }

        public int Execute(IVisio.Page page, IVisio.VisGetSetArgs flags)
        {
            if (this.items.Count < 1)
            {
                return 0;
            }

            short xflags = (short)(flags | IVisio.VisGetSetArgs.visSetUniversalSyntax);
            short[] SID_SRCStream = new short[items.Count* 4];
            object[] formulas_objects = new object[items.Count];
            for (int i = 0; i < items.Count; i++)
            {
                SID_SRCStream[i * 4 + 0] = items[i].id;
                SID_SRCStream[i * 4 + 1] = items[i].src.Section;
                SID_SRCStream[i * 4 + 2] = items[i].src.Row;
                SID_SRCStream[i * 4 + 3] = items[i].src.Cell;
                formulas_objects[i] = items[i].formula;
            }

            int count = page.SetFormulas(SID_SRCStream, formulas_objects, xflags);
            return count;
        }
    }

}