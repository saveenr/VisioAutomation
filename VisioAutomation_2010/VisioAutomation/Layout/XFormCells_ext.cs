using VisioAutomation.Drawing;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layout
{
    public partial class XFormCells
    {
        internal Rectangle GetRectangle()
        {
            var pin = new VA.Drawing.Point(this.PinX.Result, this.PinY.Result);
            var locpin = new VA.Drawing.Point(this.LocPinX.Result, this.LocPinY.Result);
            var size = new VA.Drawing.Size(this.Width.Result, this.Height.Result);
            return new VA.Drawing.Rectangle(pin - locpin, size);
        }
    }
}