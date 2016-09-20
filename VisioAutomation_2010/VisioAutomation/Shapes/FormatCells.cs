using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.CellGroups.Queries;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Shapes
{
    public class FormatCells : ShapeSheet.CellGroups.CellGroupSingleRow
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

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.FillBkgnd, this.FillBkgnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.FillBkgndTrans, this.FillBkgndTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.FillForegnd, this.FillForegnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.FillForegndTrans, this.FillForegndTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.FillPattern, this.FillPattern.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, this.ShapeShdwObliqueAngle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeShdwOffsetX, this.ShapeShdwOffsetX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeShdwOffsetY, this.ShapeShdwOffsetY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeShdwScaleFactor, this.ShapeShdwScaleFactor.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeShdwType, this.ShapeShdwType.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwBkgnd, this.ShdwBkgnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwBkgndTrans, this.ShdwBkgndTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwForegnd, this.ShdwForegnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwForegndTrans, this.ShdwForegndTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwPattern, this.ShdwPattern.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.BeginArrow, this.BeginArrow.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.BeginArrowSize, this.BeginArrowSize.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.EndArrow, this.EndArrow.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.EndArrowSize, this.EndArrowSize.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineCap, this.LineCap.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineColor, this.LineColor.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineColorTrans, this.LineColorTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LinePattern, this.LinePattern.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineWeight, this.LineWeight.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Rounding, this.Rounding.Formula);
            }
        }


        public static IList<FormatCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = FormatCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static FormatCells GetCells(IVisio.Shape shape)
        {
            var query = FormatCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static System.Lazy<ShapeFormatCellsQuery> lazy_query = new System.Lazy<ShapeFormatCellsQuery>();

    }
}

