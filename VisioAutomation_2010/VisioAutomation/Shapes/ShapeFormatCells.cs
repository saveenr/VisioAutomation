using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

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
                yield return this.newpair(ShapeSheet.SrcConstants.FillBkgnd, this.FillBkgnd.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillBkgndTrans, this.FillBkgndTrans.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillForegnd, this.FillForegnd.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillForegndTrans, this.FillForegndTrans.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.FillPattern, this.FillPattern.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeShdwObliqueAngle, this.ShapeShdwObliqueAngle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeShdwOffsetX, this.ShapeShdwOffsetX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeShdwOffsetY, this.ShapeShdwOffsetY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeShdwScaleFactor, this.ShapeShdwScaleFactor.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeShdwType, this.ShapeShdwType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShdwBkgnd, this.ShdwBkgnd.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShdwBkgndTrans, this.ShdwBkgndTrans.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShdwForegnd, this.ShdwForegnd.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShdwForegndTrans, this.ShdwForegndTrans.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShdwPattern, this.ShdwPattern.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.BeginArrow, this.BeginArrow.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.BeginArrowSize, this.BeginArrowSize.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.EndArrow, this.EndArrow.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.EndArrowSize, this.EndArrowSize.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineCap, this.LineCap.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineColor, this.LineColor.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineColorTrans, this.LineColorTrans.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LinePattern, this.LinePattern.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineWeight, this.LineWeight.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.Rounding, this.Rounding.Formula);
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

