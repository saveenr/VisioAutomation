using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageFormatCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral DrawingScale { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DrawingScaleType { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DrawingSizeType { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral InhibitSnap { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Height { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Scale { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Width { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShadowObliqueAngle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShadowOffsetX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShadowOffsetY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShadowScaleFactor { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShadowType { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral UIVisibility { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            { 
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingScale, this.DrawingScale.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingScaleType, this.DrawingScaleType.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingSizeType, this.DrawingSizeType.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageInhibitSnap, this.InhibitSnap.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageHeight, this.Height.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageScale, this.Scale.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageWidth, this.Width.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowObliqueAngle, this.ShadowObliqueAngle.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowOffsetX, this.ShadowOffsetX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowOffsetY, this.ShadowOffsetY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowScaleFactor, this.ShadowScaleFactor.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowType, this.ShadowType.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageUIVisibility, this.UIVisibility.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingResizeType, this.DrawingResizeType.Value);
            }
        }

        public static PageFormatCells GetFormulas(IVisio.Shape shape)
        {
            var query = PageFormatCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static PageFormatCells GetResults(IVisio.Shape shape)
        {
            var query = PageFormatCells.lazy_query.Value;
            return query.GetResults(shape);
        }
        private static readonly System.Lazy<PageFormatCellsReader> lazy_query = new System.Lazy<PageFormatCellsReader>();
    }
}