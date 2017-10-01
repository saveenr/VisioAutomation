using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Models
{
    public class PageFormatCells : VisioPowerShell.Models.BaseCells
    {

        // Page Format
        public string PageDrawingScale;
        public string PageDrawingScaleType;
        public string PageDrawingSizeType;
        public string PageHeight;
        public string PageScale;
        public string PageWidth;
        public string PageShadowObliqueAngle;
        public string PageShadowOffsetX;
        public string PageShadowOffsetY;
        public string PageShadowScaleFactor;
        public string PageShadowType;
        public string UIVisibility;
        public string PageDrawingResizeType;
        public string PageInhibitSnap;


        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SrcConstants.PageDrawingResizeType), SrcConstants.PageDrawingResizeType, this.PageDrawingResizeType);
            yield return new CellTuple(nameof(SrcConstants.PageDrawingScale), SrcConstants.PageDrawingScale, this.PageDrawingScale);
            yield return new CellTuple(nameof(SrcConstants.PageDrawingScaleType), SrcConstants.PageDrawingScaleType, this.PageDrawingScaleType);
            yield return new CellTuple(nameof(SrcConstants.PageDrawingSizeType), SrcConstants.PageDrawingSizeType, this.PageDrawingSizeType);
            yield return new CellTuple(nameof(SrcConstants.PageHeight), SrcConstants.PageHeight, this.PageHeight);
            yield return new CellTuple(nameof(SrcConstants.PageInhibitSnap), SrcConstants.PageInhibitSnap, this.PageInhibitSnap);
            yield return new CellTuple(nameof(SrcConstants.PageWidth), SrcConstants.PageWidth, this.PageWidth);
            yield return new CellTuple(nameof(SrcConstants.PageScale), SrcConstants.PageScale, this.PageScale);
            yield return new CellTuple(nameof(SrcConstants.PageScale), SrcConstants.PageShadowType, this.PageShadowType);
            yield return new CellTuple(nameof(SrcConstants.PageShadowObliqueAngle), SrcConstants.PageShadowObliqueAngle, this.PageShadowObliqueAngle);
            yield return new CellTuple(nameof(SrcConstants.PageShadowOffsetX), SrcConstants.PageShadowOffsetX, this.PageShadowOffsetX);
            yield return new CellTuple(nameof(SrcConstants.PageShadowOffsetY), SrcConstants.PageShadowOffsetY, this.PageShadowOffsetY);
            yield return new CellTuple(nameof(SrcConstants.PageShadowScaleFactor), SrcConstants.PageShadowScaleFactor, this.PageShadowScaleFactor);
            yield return new CellTuple(nameof(SrcConstants.PageInhibitSnap), SrcConstants.PageInhibitSnap, this.PageInhibitSnap);
        }
    }
}