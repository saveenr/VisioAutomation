using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ConnectionTests : VisioAutomationTest
    {
        public static VA.Drawing.Point GetPointAtRadius(VA.Drawing.Point origin, double angle, double radius)
        {
            var new_point = new VA.Drawing.Point(radius * System.Math.Cos(angle),
                                      radius * System.Math.Sin(angle));
            new_point = origin + new_point;
            return new_point;
        }

        [TestMethod]
        public void ConnectShapes()
        {
            var pagesize = new VA.Drawing.Size(10, 10);
            var page1 = GetNewPage(pagesize);
            var center = new VA.Drawing.Point(5, 5);

            int num = 6;
            double radius = 3;
            foreach (var i in Enumerable.Range(0, num))
            {
                double theta = i*(2*System.Math.PI)/num;
                var p = GetPointAtRadius(center, theta, radius);
                var rect = VA.Drawing.Rectangle.FromCenterPoint(p, 0.5, 0.5);
                page1.DrawOval(rect);
            }

            m_connect(page1);
            page1.Delete(0);
        }

        private static void m_connect(IVisio.Page page)
        {
            var update = new VA.ShapeSheet.Update.SRCUpdate();
            update.SetFormula(VA.ShapeSheet.SRCConstants.RouteStyle, 16);// Set page routing style to center-to-center:
            update.SetFormula(VA.ShapeSheet.SRCConstants.LineJumpStyle, 2);// Set to connector intersection to 'gap':
            var page_sheet = page.PageSheet;
            update.Execute(page_sheet);

            System.Func<IVisio.Shape, bool> is_desired_shape = shape =>
                                                               (shape.OneD == 0) &&
                                                               (shape.Type !=
                                                                (short) IVisio.VisShapeTypes.visTypeForeignObject) &&
                                                               (shape.Type != (short) IVisio.VisShapeTypes.visTypeGuide);

            var shapes_to_connect = page.Shapes.AsEnumerable()
                .Where(is_desired_shape)
                .ToList();

            var query_pairs =
                from from_shape in shapes_to_connect
                from to_shape in shapes_to_connect
                where from_shape.ID != to_shape.ID
                select new {from_shape, to_shape};

            // Connect the pairs
            foreach (var pair in query_pairs)
            {
                pair.from_shape.AutoConnect(pair.to_shape, (short) IVisio.VisAutoConnectDir.visAutoConnectDirNone, null);
            }
        }
    }
}