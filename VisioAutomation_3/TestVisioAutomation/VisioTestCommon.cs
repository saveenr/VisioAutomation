using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace TestVisioAutomation
{
    public static class VisioTestCommon
    {
        public static TestHelper Helper = new TestHelper("VisioAutomation Test Output");

        public static void SetFormulasU<T>(
            IVisio.Shape shape,
            IEnumerable<T> items,
            System.Func<T, bool> hasformula,
            System.Func<T, VA.ShapeSheet.SRC> getsrc,
            System.Func<T, string> getformula,
            IVisio.VisGetSetArgs flags )
        {
            if (items == null)
            {
                throw new System.ArgumentNullException("items");
            }

            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            update.BlastGuards = ((short)flags & (short)IVisio.VisGetSetArgs.visSetBlastGuards)!=0;
            update.TestCircular = ((short)flags & (short)IVisio.VisGetSetArgs.visSetTestCircular) != 0;

            foreach (var item in items.Where(hasformula))
            {
                update.SetFormula(getsrc(item), getformula(item));
            }

            update.Execute(shape);
        }

        public static VA.ShapeSheet.Query.CellQuery BuildCellQuery(IList<VA.ShapeSheet.SRC> srcs)
        {
            var query = new VA.ShapeSheet.Query.CellQuery();
            foreach (var src in srcs)
            {
                query.AddColumn(src);
            }
            return query;
        }

        public static void setformulas(VA.DOM.ShapeCells shapecells, IVisio.Page page, IVisio.Shape shape)
        {
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            shapecells.Apply(update,shape.ID16);
            update.Execute(page);

        }
    }
}