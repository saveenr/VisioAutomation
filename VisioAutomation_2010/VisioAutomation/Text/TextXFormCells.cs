using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;

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

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.PinX), SrcConstants.TextXFormPinX, this.PinX);
                yield return CellMetadataItem.Create(nameof(this.PinY), SrcConstants.TextXFormPinY, this.PinY);
                yield return CellMetadataItem.Create(nameof(this.LocPinX), SrcConstants.TextXFormLocPinX, this.LocPinX);
                yield return CellMetadataItem.Create(nameof(this.LocPinY), SrcConstants.TextXFormLocPinY, this.LocPinY);
                yield return CellMetadataItem.Create(nameof(this.Width), SrcConstants.TextXFormWidth, this.Width);
                yield return CellMetadataItem.Create(nameof(this.Height), SrcConstants.TextXFormHeight, this.Height);
                yield return CellMetadataItem.Create(nameof(this.Angle), SrcConstants.TextXFormAngle, this.Angle);
            }
        }
    }
}