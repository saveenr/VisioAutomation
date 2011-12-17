using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace TestVisioAutomation
{
    public class TestHelper
    {

        public static void AreEqual(double x, double y, VA.Drawing.Point p, double delta)
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(x, p.X, delta);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(y, p.Y, delta);
        }

        public static void AreEqual(double x, double y, VA.Drawing.Size p, double delta)
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(x, p.Width, delta);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(y, p.Height, delta);
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
            shapecells.Apply(update, shape.ID16);
            update.Execute(page);
        }

    }
}