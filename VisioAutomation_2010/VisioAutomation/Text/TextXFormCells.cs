using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextXFormCells : CellGroup
    {
        public CellValueLiteral Angle { get; set; }
        public CellValueLiteral Width { get; set; }
        public CellValueLiteral Height { get; set; }
        public CellValueLiteral PinX { get; set; }
        public CellValueLiteral PinY { get; set; }
        public CellValueLiteral LocPinX { get; set; }
        public CellValueLiteral LocPinY { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.TextXFormPinX, this.PinX);
                yield return SrcValuePair.Create(SrcConstants.TextXFormPinY, this.PinY);
                yield return SrcValuePair.Create(SrcConstants.TextXFormLocPinX, this.LocPinX);
                yield return SrcValuePair.Create(SrcConstants.TextXFormLocPinY, this.LocPinY);
                yield return SrcValuePair.Create(SrcConstants.TextXFormWidth, this.Width);
                yield return SrcValuePair.Create(SrcConstants.TextXFormHeight, this.Height);
                yield return SrcValuePair.Create(SrcConstants.TextXFormAngle, this.Angle);
            }
        }
    }
}