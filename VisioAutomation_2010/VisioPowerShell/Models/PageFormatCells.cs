using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Models
{
    public class PageFormatCells : VisioPowerShell.Models.BaseCells
    {
        public string DrawingScale;
        public string DrawingScaleType;
        public string DrawingSizeType;
        public string Height;

        public string Scale;
        public string Width;

        public string ShadowObliqueAngle;
        public string ShadowOffsetX;
        public string ShadowOffsetY;
        public string ShadowScaleFactor;
        public string ShadowType;
        public string UIVisibility;
        public string DrawingResizeType;
        public string InhibitSnap;


        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SrcConstants.PageDrawingResizeType), SrcConstants.PageDrawingResizeType, this.DrawingResizeType);
            yield return new CellTuple(nameof(SrcConstants.PageDrawingScale), SrcConstants.PageDrawingScale, this.DrawingScale);
            yield return new CellTuple(nameof(SrcConstants.PageDrawingScaleType), SrcConstants.PageDrawingScaleType, this.DrawingScaleType);
            yield return new CellTuple(nameof(SrcConstants.PageDrawingSizeType), SrcConstants.PageDrawingSizeType, this.DrawingSizeType);
            yield return new CellTuple(nameof(SrcConstants.PageHeight), SrcConstants.PageHeight, this.Height);
            yield return new CellTuple(nameof(SrcConstants.PageInhibitSnap), SrcConstants.PageInhibitSnap, this.InhibitSnap);
            yield return new CellTuple(nameof(SrcConstants.PageWidth), SrcConstants.PageWidth, this.Width);
            yield return new CellTuple(nameof(SrcConstants.PageScale), SrcConstants.PageScale, this.Scale);

            yield return new CellTuple(nameof(SrcConstants.PageScale), SrcConstants.PageShadowType, this.ShadowType);
            yield return new CellTuple(nameof(SrcConstants.PageScale), SrcConstants.PageShadowObliqueAngle, this.ShadowObliqueAngle);
            yield return new CellTuple(nameof(SrcConstants.PageScale), SrcConstants.PageShadowOffsetX, this.ShadowOffsetX);
            yield return new CellTuple(nameof(SrcConstants.PageScale), SrcConstants.PageShadowOffsetY, this.ShadowOffsetY);
            yield return new CellTuple(nameof(SrcConstants.PageScale), SrcConstants.PageShadowScaleFactor, this.ShadowScaleFactor);
            yield return new CellTuple(nameof(SrcConstants.PageInhibitSnap), SrcConstants.PageInhibitSnap, this.InhibitSnap);
        }
    }
}