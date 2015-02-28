using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class FormatCells : VA.ShapeSheet.CellGroups.CellGroup
    {
        public VA.ShapeSheet.CellData<int> FillBkgnd { get; set; }
        public VA.ShapeSheet.CellData<double> FillBkgndTrans { get; set; }
        public VA.ShapeSheet.CellData<int> FillForegnd { get; set; }
        public VA.ShapeSheet.CellData<double> FillForegndTrans { get; set; }
        public VA.ShapeSheet.CellData<int> FillPattern { get; set; }
        public VA.ShapeSheet.CellData<double> ShapeShdwObliqueAngle { get; set; }
        public VA.ShapeSheet.CellData<double> ShapeShdwOffsetX { get; set; }
        public VA.ShapeSheet.CellData<double> ShapeShdwOffsetY { get; set; }
        public VA.ShapeSheet.CellData<double> ShapeShdwScaleFactor { get; set; }
        public VA.ShapeSheet.CellData<int> ShapeShdwType { get; set; }
        public VA.ShapeSheet.CellData<int> ShdwBkgnd { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwBkgndTrans { get; set; }
        public VA.ShapeSheet.CellData<int> ShdwForegnd { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwForegndTrans { get; set; }
        public VA.ShapeSheet.CellData<int> ShdwPattern { get; set; }
        public VA.ShapeSheet.CellData<int> BeginArrow { get; set; }
        public VA.ShapeSheet.CellData<double> BeginArrowSize { get; set; }
        public VA.ShapeSheet.CellData<int> EndArrow { get; set; }
        public VA.ShapeSheet.CellData<double> EndArrowSize { get; set; }
        public VA.ShapeSheet.CellData<int> LineCap { get; set; }
        public VA.ShapeSheet.CellData<int> LineColor { get; set; }
        public VA.ShapeSheet.CellData<double> LineColorTrans { get; set; }
        public VA.ShapeSheet.CellData<int> LinePattern { get; set; }
        public VA.ShapeSheet.CellData<double> LineWeight { get; set; }
        public VA.ShapeSheet.CellData<double> Rounding { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return newpair(ShapeSheet.SRCConstants.FillBkgnd, this.FillBkgnd.Formula);
                yield return newpair(ShapeSheet.SRCConstants.FillBkgndTrans, this.FillBkgndTrans.Formula);
                yield return newpair(ShapeSheet.SRCConstants.FillForegnd, this.FillForegnd.Formula);
                yield return newpair(ShapeSheet.SRCConstants.FillForegndTrans, this.FillForegndTrans.Formula);
                yield return newpair(ShapeSheet.SRCConstants.FillPattern, this.FillPattern.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, this.ShapeShdwObliqueAngle.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapeShdwOffsetX, this.ShapeShdwOffsetX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapeShdwOffsetY, this.ShapeShdwOffsetY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapeShdwScaleFactor, this.ShapeShdwScaleFactor.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShapeShdwType, this.ShapeShdwType.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShdwBkgnd, this.ShdwBkgnd.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShdwBkgndTrans, this.ShdwBkgndTrans.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShdwForegnd, this.ShdwForegnd.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShdwForegndTrans, this.ShdwForegndTrans.Formula);
                yield return newpair(ShapeSheet.SRCConstants.ShdwPattern, this.ShdwPattern.Formula);
                yield return newpair(ShapeSheet.SRCConstants.BeginArrow, this.BeginArrow.Formula);
                yield return newpair(ShapeSheet.SRCConstants.BeginArrowSize, this.BeginArrowSize.Formula);
                yield return newpair(ShapeSheet.SRCConstants.EndArrow, this.EndArrow.Formula);
                yield return newpair(ShapeSheet.SRCConstants.EndArrowSize, this.EndArrowSize.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineCap, this.LineCap.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineColor, this.LineColor.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineColorTrans, this.LineColorTrans.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LinePattern, this.LinePattern.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LineWeight, this.LineWeight.Formula);
                yield return newpair(ShapeSheet.SRCConstants.Rounding, this.Rounding.Formula);
            }
        }


        public static IList<FormatCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells<FormatCells,double>(page, shapeids, query, query.GetCells);
        }

        public static FormatCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells<FormatCells, double>(shape, query, query.GetCells);
        }

        private static ShapeFormatCellQuery _mCellQuery;
        private static ShapeFormatCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new ShapeFormatCellQuery();
            return _mCellQuery;
        }

        class ShapeFormatCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public CellColumn FillBkgnd { get; set; }
            public CellColumn FillBkgndTrans { get; set; }
            public CellColumn FillForegnd { get; set; }
            public CellColumn FillForegndTrans { get; set; }
            public CellColumn FillPattern { get; set; }
            public CellColumn ShapeShdwObliqueAngle { get; set; }
            public CellColumn ShapeShdwOffsetX { get; set; }
            public CellColumn ShapeShdwOffsetY { get; set; }
            public CellColumn ShapeShdwScaleFactor { get; set; }
            public CellColumn ShapeShdwType { get; set; }
            public CellColumn ShdwBkgnd { get; set; }
            public CellColumn ShdwBkgndTrans { get; set; }
            public CellColumn ShdwForegnd { get; set; }
            public CellColumn ShdwForegndTrans { get; set; }
            public CellColumn ShdwPattern { get; set; }
            public CellColumn BeginArrow { get; set; }
            public CellColumn BeginArrowSize { get; set; }
            public CellColumn EndArrow { get; set; }
            public CellColumn EndArrowSize { get; set; }
            public CellColumn LineColor { get; set; }
            public CellColumn LineCap { get; set; }
            public CellColumn LineColorTrans { get; set; }
            public CellColumn LinePattern { get; set; }
            public CellColumn LineWeight { get; set; }
            public CellColumn Rounding { get; set; }

            public ShapeFormatCellQuery()
            {
                this.FillBkgnd = this.AddCell(VA.ShapeSheet.SRCConstants.FillBkgnd);
                this.FillBkgndTrans = this.AddCell(VA.ShapeSheet.SRCConstants.FillBkgndTrans);
                this.FillForegnd = this.AddCell(VA.ShapeSheet.SRCConstants.FillForegnd);
                this.FillForegndTrans = this.AddCell(VA.ShapeSheet.SRCConstants.FillForegndTrans);
                this.FillPattern = this.AddCell(VA.ShapeSheet.SRCConstants.FillPattern);
                this.ShapeShdwObliqueAngle = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeShdwObliqueAngle);
                this.ShapeShdwOffsetX = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeShdwOffsetX);
                this.ShapeShdwOffsetY = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeShdwOffsetY);
                this.ShapeShdwScaleFactor = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeShdwScaleFactor);
                this.ShapeShdwType = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeShdwType);
                this.ShdwBkgnd = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwBkgnd);
                this.ShdwBkgndTrans = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwBkgndTrans);
                this.ShdwForegnd = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwForegnd);
                this.ShdwForegndTrans = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwForegndTrans);
                this.ShdwPattern = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwPattern);

                this.BeginArrow = this.AddCell(VA.ShapeSheet.SRCConstants.BeginArrow);
                this.BeginArrowSize = this.AddCell(VA.ShapeSheet.SRCConstants.BeginArrowSize);
                this.EndArrow = this.AddCell(VA.ShapeSheet.SRCConstants.EndArrow);
                this.EndArrowSize = this.AddCell(VA.ShapeSheet.SRCConstants.EndArrowSize);
                this.LineColor = this.AddCell(VA.ShapeSheet.SRCConstants.LineColor);
                this.LineCap = this.AddCell(VA.ShapeSheet.SRCConstants.LineCap);
                this.LineColorTrans = this.AddCell(VA.ShapeSheet.SRCConstants.LineColorTrans);
                this.LinePattern = this.AddCell(VA.ShapeSheet.SRCConstants.LinePattern);
                this.LineWeight = this.AddCell(VA.ShapeSheet.SRCConstants.LineWeight);
                this.Rounding = this.AddCell(VA.ShapeSheet.SRCConstants.Rounding);
            }

            public FormatCells GetCells(IList<VA.ShapeSheet.CellData<double>> row)
            {
                var cells = new FormatCells();
                cells.FillBkgnd = row[ this.FillBkgnd].ToInt();
                cells.FillBkgndTrans = row[ this.FillBkgndTrans];
                cells.FillForegnd = row[ this.FillForegnd].ToInt();
                cells.FillForegndTrans = row[ this.FillForegndTrans];
                cells.FillPattern = row[ this.FillPattern].ToInt();
                cells.ShapeShdwObliqueAngle = row[ this.ShapeShdwObliqueAngle];
                cells.ShapeShdwOffsetX = row[ this.ShapeShdwOffsetX];
                cells.ShapeShdwOffsetY = row[ this.ShapeShdwOffsetY];
                cells.ShapeShdwScaleFactor = row[ this.ShapeShdwScaleFactor];
                cells.ShapeShdwType = row[ this.ShapeShdwType].ToInt();
                cells.ShdwBkgnd = row[ this.ShdwBkgnd].ToInt();
                cells.ShdwBkgndTrans = row[ this.ShdwBkgndTrans];
                cells.ShdwForegnd = row[ this.ShdwForegnd].ToInt();
                cells.ShdwForegndTrans = row[ this.ShdwForegndTrans];
                cells.ShdwPattern = row[ this.ShdwPattern].ToInt();
                cells.BeginArrow = row[ this.BeginArrow].ToInt();
                cells.BeginArrowSize = row[ this.BeginArrowSize];
                cells.EndArrow = row[ this.EndArrow].ToInt();
                cells.EndArrowSize = row[ this.EndArrowSize];
                cells.LineCap = row[ this.LineCap].ToInt();
                cells.LineColor = row[ this.LineColor].ToInt();
                cells.LineColorTrans = row[ this.LineColorTrans];
                cells.LinePattern = row[ this.LinePattern].ToInt();
                cells.LineWeight = row[ this.LineWeight];
                cells.Rounding = row[ this.Rounding];
                return cells;
            }

        }
    }
}