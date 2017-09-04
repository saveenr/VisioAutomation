using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ShapeFormatCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData FillBackground { get; set; }
        public ShapeSheet.CellData FillBackgroundTransparency { get; set; }
        public ShapeSheet.CellData FillForeground { get; set; }
        public ShapeSheet.CellData FillForegroundTransparency { get; set; }
        public ShapeSheet.CellData FillPattern { get; set; }
        public ShapeSheet.CellData FillShadowObliqueAngle { get; set; }
        public ShapeSheet.CellData FillShadowOffsetX { get; set; }
        public ShapeSheet.CellData FillShadowOffsetY { get; set; }
        public ShapeSheet.CellData FillShadowScaleFactor { get; set; }
        public ShapeSheet.CellData FillShadowType { get; set; }
        public ShapeSheet.CellData FillShadowBackground { get; set; }
        public ShapeSheet.CellData FillShadowBackgroundTransparency { get; set; }
        public ShapeSheet.CellData FillShadowForeground { get; set; }
        public ShapeSheet.CellData FillShadowForegroundTransparency { get; set; }
        public ShapeSheet.CellData FillShadowPattern { get; set; }
        public ShapeSheet.CellData LineBeginArrow { get; set; }
        public ShapeSheet.CellData LineBeginArrowSize { get; set; }
        public ShapeSheet.CellData LineEndArrow { get; set; }
        public ShapeSheet.CellData LineEndArrowSize { get; set; }
        public ShapeSheet.CellData LineCap { get; set; }
        public ShapeSheet.CellData LineColor { get; set; }
        public ShapeSheet.CellData LineColorTransparency { get; set; }
        public ShapeSheet.CellData LinePattern { get; set; }
        public ShapeSheet.CellData LineWeight { get; set; }
        public ShapeSheet.CellData LineRounding { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.FillBackground, this.FillBackground.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillBackgroundTransparency, this.FillBackgroundTransparency.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillForeground, this.FillForeground.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillForegroundTransparency, this.FillForegroundTransparency.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillPattern, this.FillPattern.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowObliqueAngle, this.FillShadowObliqueAngle.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowOffsetX, this.FillShadowOffsetX.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowOffsetY, this.FillShadowOffsetY.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowScaleFactor, this.FillShadowScaleFactor.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowType, this.FillShadowType.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowBackground, this.FillShadowBackground.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowForeground, this.FillShadowForeground.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowPattern, this.FillShadowPattern.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LineBeginArrow, this.LineBeginArrow.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LineBeginArrowSize, this.LineBeginArrowSize.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LineEndArrow, this.LineEndArrow.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LineEndArrowSize, this.LineEndArrowSize.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LineCap, this.LineCap.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LineColor, this.LineColor.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LineColorTransparency, this.LineColorTransparency.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LinePattern, this.LinePattern.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LineWeight, this.LineWeight.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LineRounding, this.LineRounding.ValueF);
            }
        }


        public static List<ShapeFormatCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = ShapeFormatCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static ShapeFormatCells GetCells(IVisio.Shape shape)
        {
            var query = ShapeFormatCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<ShapeFormatCellsReader> lazy_query = new System.Lazy<ShapeFormatCellsReader>();
    }
}

