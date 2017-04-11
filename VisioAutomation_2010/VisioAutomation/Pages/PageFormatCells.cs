using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageFormatCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData PageDrawingScale { get; set; }
        public ShapeSheet.CellData PageDrawingScaleType { get; set; }
        public ShapeSheet.CellData PageDrawingSizeType { get; set; }
        public ShapeSheet.CellData PageInhibitSnap { get; set; }
        public ShapeSheet.CellData PageHeight { get; set; }
        public ShapeSheet.CellData PageScale { get; set; }
        public ShapeSheet.CellData PageWidth { get; set; }
        public ShapeSheet.CellData PageShadowObliqueAngle { get; set; }
        public ShapeSheet.CellData PageShadowOffsetX { get; set; }
        public ShapeSheet.CellData PageShadowOffsetY { get; set; }
        public ShapeSheet.CellData PageShadowScaleFactor { get; set; }
        public ShapeSheet.CellData PageShadowType { get; set; }
        public ShapeSheet.CellData PageUIVisibility { get; set; }
        public ShapeSheet.CellData PageDrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            { 
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingScale, this.PageDrawingScale.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingScaleType, this.PageDrawingScaleType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingSizeType, this.PageDrawingSizeType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageInhibitSnap, this.PageInhibitSnap.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageHeight, this.PageHeight.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageScale, this.PageScale.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageWidth, this.PageWidth.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowObliqueAngle, this.PageShadowObliqueAngle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowOffsetX, this.PageShadowOffsetX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowOffsetY, this.PageShadowOffsetY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowScaleFactor, this.PageShadowScaleFactor.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowType, this.PageShadowType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageUIVisibility, this.PageUIVisibility.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingResizeType, this.PageDrawingResizeType.Formula);
            }
        }

        public static PageFormatCells GetCells(IVisio.Shape shape)
        {
            var query = PageFormatCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<PageFormatCellsReader> lazy_query = new System.Lazy<PageFormatCellsReader>();
    }
}