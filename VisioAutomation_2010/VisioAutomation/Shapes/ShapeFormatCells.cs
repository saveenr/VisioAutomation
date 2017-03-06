using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ShapeFormatCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData FillBkgnd { get; set; }
        public ShapeSheet.CellData FillBkgndTrans { get; set; }
        public ShapeSheet.CellData FillForegnd { get; set; }
        public ShapeSheet.CellData FillForegndTrans { get; set; }
        public ShapeSheet.CellData FillPattern { get; set; }
        public ShapeSheet.CellData ShapeShdwObliqueAngle { get; set; }
        public ShapeSheet.CellData ShapeShdwOffsetX { get; set; }
        public ShapeSheet.CellData ShapeShdwOffsetY { get; set; }
        public ShapeSheet.CellData ShapeShdwScaleFactor { get; set; }
        public ShapeSheet.CellData ShapeShdwType { get; set; }
        public ShapeSheet.CellData ShdwBkgnd { get; set; }
        public ShapeSheet.CellData ShdwBkgndTrans { get; set; }
        public ShapeSheet.CellData ShdwForegnd { get; set; }
        public ShapeSheet.CellData ShdwForegndTrans { get; set; }
        public ShapeSheet.CellData ShdwPattern { get; set; }
        public ShapeSheet.CellData BeginArrow { get; set; }
        public ShapeSheet.CellData BeginArrowSize { get; set; }
        public ShapeSheet.CellData EndArrow { get; set; }
        public ShapeSheet.CellData EndArrowSize { get; set; }
        public ShapeSheet.CellData LineCap { get; set; }
        public ShapeSheet.CellData LineColor { get; set; }
        public ShapeSheet.CellData LineColorTrans { get; set; }
        public ShapeSheet.CellData LinePattern { get; set; }
        public ShapeSheet.CellData LineWeight { get; set; }
        public ShapeSheet.CellData Rounding { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.FillBackground, this.FillBkgnd.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillBackgroundTransparency, this.FillBkgndTrans.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillForeground, this.FillForegnd.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillForegroundTransparency, this.FillForegndTrans.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillPattern, this.FillPattern.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowObliqueAngle, this.ShapeShdwObliqueAngle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowOffsetX, this.ShapeShdwOffsetX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowOffsetY, this.ShapeShdwOffsetY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowScaleFactor, this.ShapeShdwScaleFactor.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowType, this.ShapeShdwType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowBackground, this.ShdwBkgnd.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowBackgroundTransparency, this.ShdwBkgndTrans.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowForeground, this.ShdwForegnd.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowForegroundTransparency, this.ShdwForegndTrans.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillShadowPattern, this.ShdwPattern.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineBeginArrow, this.BeginArrow.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineBeginArrowSize, this.BeginArrowSize.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineEndArrow, this.EndArrow.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineEndArrowSize, this.EndArrowSize.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineCap, this.LineCap.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineColor, this.LineColor.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineColorTransparency, this.LineColorTrans.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LinePattern, this.LinePattern.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineWeight, this.LineWeight.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineRounding, this.Rounding.Formula);
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

