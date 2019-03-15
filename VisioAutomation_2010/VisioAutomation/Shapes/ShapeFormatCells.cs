using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public class ShapeFormatCells : CellGroup
    {
        public CellValueLiteral FillBackground { get; set; }
        public CellValueLiteral FillBackgroundTransparency { get; set; }
        public CellValueLiteral FillForeground { get; set; }
        public CellValueLiteral FillForegroundTransparency { get; set; }
        public CellValueLiteral FillPattern { get; set; }
        public CellValueLiteral FillShadowObliqueAngle { get; set; }
        public CellValueLiteral FillShadowOffsetX { get; set; }
        public CellValueLiteral FillShadowOffsetY { get; set; }
        public CellValueLiteral FillShadowScaleFactor { get; set; }
        public CellValueLiteral FillShadowType { get; set; }
        public CellValueLiteral FillShadowBackground { get; set; }
        public CellValueLiteral FillShadowBackgroundTransparency { get; set; }
        public CellValueLiteral FillShadowForeground { get; set; }
        public CellValueLiteral FillShadowForegroundTransparency { get; set; }
        public CellValueLiteral FillShadowPattern { get; set; }
        public CellValueLiteral LineBeginArrow { get; set; }
        public CellValueLiteral LineBeginArrowSize { get; set; }
        public CellValueLiteral LineEndArrow { get; set; }
        public CellValueLiteral LineEndArrowSize { get; set; }
        public CellValueLiteral LineCap { get; set; }
        public CellValueLiteral LineColor { get; set; }
        public CellValueLiteral LineColorTransparency { get; set; }
        public CellValueLiteral LinePattern { get; set; }
        public CellValueLiteral LineWeight { get; set; }
        public CellValueLiteral LineRounding { get; set; }

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.FillBackground), SrcConstants.FillBackground, this.FillBackground);
                yield return CellMetadataItem.Create(nameof(this.FillBackgroundTransparency), SrcConstants.FillBackgroundTransparency, this.FillBackgroundTransparency);
                yield return CellMetadataItem.Create(nameof(this.FillForeground), SrcConstants.FillForeground, this.FillForeground);
                yield return CellMetadataItem.Create(nameof(this.FillForegroundTransparency), SrcConstants.FillForegroundTransparency, this.FillForegroundTransparency);
                yield return CellMetadataItem.Create(nameof(this.FillPattern), SrcConstants.FillPattern, this.FillPattern);
                yield return CellMetadataItem.Create(nameof(this.FillShadowObliqueAngle), SrcConstants.FillShadowObliqueAngle, this.FillShadowObliqueAngle);
                yield return CellMetadataItem.Create(nameof(this.FillShadowOffsetX), SrcConstants.FillShadowOffsetX, this.FillShadowOffsetX);
                yield return CellMetadataItem.Create(nameof(this.FillShadowOffsetY), SrcConstants.FillShadowOffsetY, this.FillShadowOffsetY);
                yield return CellMetadataItem.Create(nameof(this.FillShadowScaleFactor), SrcConstants.FillShadowScaleFactor, this.FillShadowScaleFactor);
                yield return CellMetadataItem.Create(nameof(this.FillShadowType), SrcConstants.FillShadowType, this.FillShadowType);
                yield return CellMetadataItem.Create(nameof(this.FillShadowBackground), SrcConstants.FillShadowBackground, this.FillShadowBackground);
                yield return CellMetadataItem.Create(nameof(this.FillShadowBackgroundTransparency), SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
                yield return CellMetadataItem.Create(nameof(this.FillShadowForeground), SrcConstants.FillShadowForeground, this.FillShadowForeground);
                yield return CellMetadataItem.Create(nameof(this.FillShadowForegroundTransparency), SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
                yield return CellMetadataItem.Create(nameof(this.FillShadowPattern), SrcConstants.FillShadowPattern, this.FillShadowPattern);
                yield return CellMetadataItem.Create(nameof(this.LineBeginArrow), SrcConstants.LineBeginArrow, this.LineBeginArrow);
                yield return CellMetadataItem.Create(nameof(this.LineBeginArrowSize), SrcConstants.LineBeginArrowSize, this.LineBeginArrowSize);
                yield return CellMetadataItem.Create(nameof(this.LineEndArrow), SrcConstants.LineEndArrow, this.LineEndArrow);
                yield return CellMetadataItem.Create(nameof(this.LineEndArrowSize), SrcConstants.LineEndArrowSize, this.LineEndArrowSize);
                yield return CellMetadataItem.Create(nameof(this.LineCap), SrcConstants.LineCap, this.LineCap);
                yield return CellMetadataItem.Create(nameof(this.LineColor), SrcConstants.LineColor, this.LineColor);
                yield return CellMetadataItem.Create(nameof(this.LineColorTransparency), SrcConstants.LineColorTransparency, this.LineColorTransparency);
                yield return CellMetadataItem.Create(nameof(this.LinePattern), SrcConstants.LinePattern, this.LinePattern);
                yield return CellMetadataItem.Create(nameof(this.LineWeight), SrcConstants.LineWeight, this.LineWeight);
                yield return CellMetadataItem.Create(nameof(this.LineRounding), SrcConstants.LineRounding, this.LineRounding);
            }
        }
    }
}

