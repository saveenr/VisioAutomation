using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Shapes
{
    public class FormatCells : ShapeSheetQuery.QueryGroups.QueryGroupSingleRow
    {
        public ShapeSheet.CellData<int> FillBkgnd { get; set; }
        public ShapeSheet.CellData<double> FillBkgndTrans { get; set; }
        public ShapeSheet.CellData<int> FillForegnd { get; set; }
        public ShapeSheet.CellData<double> FillForegndTrans { get; set; }
        public ShapeSheet.CellData<int> FillPattern { get; set; }
        public ShapeSheet.CellData<double> ShapeShdwObliqueAngle { get; set; }
        public ShapeSheet.CellData<double> ShapeShdwOffsetX { get; set; }
        public ShapeSheet.CellData<double> ShapeShdwOffsetY { get; set; }
        public ShapeSheet.CellData<double> ShapeShdwScaleFactor { get; set; }
        public ShapeSheet.CellData<int> ShapeShdwType { get; set; }
        public ShapeSheet.CellData<int> ShdwBkgnd { get; set; }
        public ShapeSheet.CellData<double> ShdwBkgndTrans { get; set; }
        public ShapeSheet.CellData<int> ShdwForegnd { get; set; }
        public ShapeSheet.CellData<double> ShdwForegndTrans { get; set; }
        public ShapeSheet.CellData<int> ShdwPattern { get; set; }
        public ShapeSheet.CellData<int> BeginArrow { get; set; }
        public ShapeSheet.CellData<double> BeginArrowSize { get; set; }
        public ShapeSheet.CellData<int> EndArrow { get; set; }
        public ShapeSheet.CellData<double> EndArrowSize { get; set; }
        public ShapeSheet.CellData<int> LineCap { get; set; }
        public ShapeSheet.CellData<int> LineColor { get; set; }
        public ShapeSheet.CellData<double> LineColorTrans { get; set; }
        public ShapeSheet.CellData<int> LinePattern { get; set; }
        public ShapeSheet.CellData<double> LineWeight { get; set; }
        public ShapeSheet.CellData<double> Rounding { get; set; }

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
            return ShapeSheetQuery.QueryGroups.QueryGroupSingleRow._GetCells<FormatCells, double>(page, shapeids, query, query.GetCells);
        }

        public static FormatCells GetCells(IVisio.Shape shape)
        {
            var query = FormatCells.lazy_query.Value;
            return ShapeSheetQuery.QueryGroups.QueryGroupSingleRow._GetCells<FormatCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<VA.ShapeSheetQuery.Common.ShapeFormatCellsQuery> lazy_query = new System.Lazy<VA.ShapeSheetQuery.Common.ShapeFormatCellsQuery>();

    }
}

