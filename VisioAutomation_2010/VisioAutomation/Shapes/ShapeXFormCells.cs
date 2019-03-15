using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public class ShapeXFormCells : CellGroup
    {
        public CellValueLiteral PinX { get; set; }
        public CellValueLiteral PinY { get; set; }
        public CellValueLiteral LocPinX { get; set; }
        public CellValueLiteral LocPinY { get; set; }
        public CellValueLiteral Width { get; set; }
        public CellValueLiteral Height { get; set; }
        public CellValueLiteral Angle { get; set; }

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.PinX), SrcConstants.XFormPinX, this.PinX);
                yield return CellMetadataItem.Create(nameof(this.PinY), SrcConstants.XFormPinY, this.PinY);
                yield return CellMetadataItem.Create(nameof(this.LocPinX), SrcConstants.XFormLocPinX, this.LocPinX);
                yield return CellMetadataItem.Create(nameof(this.LocPinY), SrcConstants.XFormLocPinY, this.LocPinY);
                yield return CellMetadataItem.Create(nameof(this.Width), SrcConstants.XFormWidth, this.Width);
                yield return CellMetadataItem.Create(nameof(this.Height), SrcConstants.XFormHeight, this.Height);
                yield return CellMetadataItem.Create(nameof(this.Angle), SrcConstants.XFormAngle, this.Angle);
            }
        }
    }
}