using VisioAutomation.Shapes;

namespace TestVisioAutomation
{
    public static class TestExtensions
    {
        public static VisioAutomation.Drawing.Point Pin( this XFormCells xthis)
        {
            return new VisioAutomation.Drawing.Point(xthis.PinX.Result, xthis.PinY.Result);
        }
    }
}