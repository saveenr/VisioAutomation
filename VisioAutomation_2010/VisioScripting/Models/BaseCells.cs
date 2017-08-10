using System.Collections.Generic;

namespace VisioScripting.Models
{
    public struct CellTuple
    {
        public string Name;
        public VisioAutomation.ShapeSheet.Src Src;
        public string Formula;

        public CellTuple(string name, VisioAutomation.ShapeSheet.Src src, string formula)
        {
            this.Name = name;
            this.Src = src;
            this.Formula = formula;
        }
    }

    public abstract class BaseCells
    {
        public abstract IEnumerable<CellTuple> GetSrcFormulaPairs();

        public void Apply(VisioAutomation.ShapeSheet.Writers.SidSrcWriter writer, short id)
        {
            foreach (var pair in this.GetSrcFormulaPairs())
            {
                if (pair.Formula != null)
                {
                    writer.SetFormula(id, pair.Src, pair.Formula);
                }
            }
        }

    }
}