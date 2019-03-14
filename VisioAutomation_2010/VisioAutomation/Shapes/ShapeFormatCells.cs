using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

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

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.FillBackground, this.FillBackground);
                yield return SrcValuePair.Create(SrcConstants.FillBackgroundTransparency, this.FillBackgroundTransparency);
                yield return SrcValuePair.Create(SrcConstants.FillForeground, this.FillForeground);
                yield return SrcValuePair.Create(SrcConstants.FillForegroundTransparency, this.FillForegroundTransparency);
                yield return SrcValuePair.Create(SrcConstants.FillPattern, this.FillPattern);
                yield return SrcValuePair.Create(SrcConstants.FillShadowObliqueAngle, this.FillShadowObliqueAngle);
                yield return SrcValuePair.Create(SrcConstants.FillShadowOffsetX, this.FillShadowOffsetX);
                yield return SrcValuePair.Create(SrcConstants.FillShadowOffsetY, this.FillShadowOffsetY);
                yield return SrcValuePair.Create(SrcConstants.FillShadowScaleFactor, this.FillShadowScaleFactor);
                yield return SrcValuePair.Create(SrcConstants.FillShadowType, this.FillShadowType);
                yield return SrcValuePair.Create(SrcConstants.FillShadowBackground, this.FillShadowBackground);
                yield return SrcValuePair.Create(SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
                yield return SrcValuePair.Create(SrcConstants.FillShadowForeground, this.FillShadowForeground);
                yield return SrcValuePair.Create(SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
                yield return SrcValuePair.Create(SrcConstants.FillShadowPattern, this.FillShadowPattern);
                yield return SrcValuePair.Create(SrcConstants.LineBeginArrow, this.LineBeginArrow);
                yield return SrcValuePair.Create(SrcConstants.LineBeginArrowSize, this.LineBeginArrowSize);
                yield return SrcValuePair.Create(SrcConstants.LineEndArrow, this.LineEndArrow);
                yield return SrcValuePair.Create(SrcConstants.LineEndArrowSize, this.LineEndArrowSize);
                yield return SrcValuePair.Create(SrcConstants.LineCap, this.LineCap);
                yield return SrcValuePair.Create(SrcConstants.LineColor, this.LineColor);
                yield return SrcValuePair.Create(SrcConstants.LineColorTransparency, this.LineColorTransparency);
                yield return SrcValuePair.Create(SrcConstants.LinePattern, this.LinePattern);
                yield return SrcValuePair.Create(SrcConstants.LineWeight, this.LineWeight);
                yield return SrcValuePair.Create(SrcConstants.LineRounding, this.LineRounding);
            }
        }
    }
}

