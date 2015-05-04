using VisioAutomation.Drawing;
using VisioAutomation.Shapes;

namespace TestVisioAutomation
{
    public static class TestExtensions
    {
        public static Point Pin(this XFormCells xform)
        {
            return new Point(xform.PinX.Result, xform.PinY.Result);
        }
    }
}