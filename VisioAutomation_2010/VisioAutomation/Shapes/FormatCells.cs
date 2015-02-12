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
                this.FillBkgnd = this.AddCell(VA.ShapeSheet.SRCConstants.FillBkgnd, "FillBkgnd");
                this.FillBkgndTrans = this.AddCell(VA.ShapeSheet.SRCConstants.FillBkgndTrans, "FillBkgndTrans");
                this.FillForegnd = this.AddCell(VA.ShapeSheet.SRCConstants.FillForegnd, "FillForegnd");
                this.FillForegndTrans = this.AddCell(VA.ShapeSheet.SRCConstants.FillForegndTrans, "FillForegndTrans");
                this.FillPattern = this.AddCell(VA.ShapeSheet.SRCConstants.FillPattern, "FillPattern");
                this.ShapeShdwObliqueAngle = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, "ShapeShdwObliqueAngle");
                this.ShapeShdwOffsetX = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeShdwOffsetX, "ShapeShdwOffsetX");
                this.ShapeShdwOffsetY = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeShdwOffsetY, "ShapeShdwOffsetY");
                this.ShapeShdwScaleFactor = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeShdwScaleFactor, "ShapeShdwScaleFactor");
                this.ShapeShdwType = this.AddCell(VA.ShapeSheet.SRCConstants.ShapeShdwType, "ShapeShdwType");
                this.ShdwBkgnd = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwBkgnd, "ShdwBkgnd ");
                this.ShdwBkgndTrans = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwBkgndTrans, "ShdwBkgndTrans");
                this.ShdwForegnd = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwForegnd, "ShdwForegnd ");
                this.ShdwForegndTrans = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwForegndTrans, "ShdwForegndTrans");
                this.ShdwPattern = this.AddCell(VA.ShapeSheet.SRCConstants.ShdwPattern, "ShdwPattern");

                this.BeginArrow = this.AddCell(VA.ShapeSheet.SRCConstants.BeginArrow, "BeginArrow");
                this.BeginArrowSize = this.AddCell(VA.ShapeSheet.SRCConstants.BeginArrowSize, "BeginArrowSize");
                this.EndArrow = this.AddCell(VA.ShapeSheet.SRCConstants.EndArrow, "EndArrow ");
                this.EndArrowSize = this.AddCell(VA.ShapeSheet.SRCConstants.EndArrowSize, "EndArrowSize");
                this.LineColor = this.AddCell(VA.ShapeSheet.SRCConstants.LineColor, "LineColor");
                this.LineCap = this.AddCell(VA.ShapeSheet.SRCConstants.LineCap, "LineCap");
                this.LineColorTrans = this.AddCell(VA.ShapeSheet.SRCConstants.LineColorTrans, "LineColorTrans");
                this.LinePattern = this.AddCell(VA.ShapeSheet.SRCConstants.LinePattern, "LinePattern");
                this.LineWeight = this.AddCell(VA.ShapeSheet.SRCConstants.LineWeight, "LineWeight");
                this.Rounding = this.AddCell(VA.ShapeSheet.SRCConstants.Rounding, "Rounding");
            }

            public FormatCells GetCells(IList<VA.ShapeSheet.CellData<double>> row)
            {
                var cells = new FormatCells();
                cells.FillBkgnd = row[ this.FillBkgnd.Ordinal].ToInt();
                cells.FillBkgndTrans = row[ this.FillBkgndTrans.Ordinal];
                cells.FillForegnd = row[ this.FillForegnd.Ordinal].ToInt();
                cells.FillForegndTrans = row[ this.FillForegndTrans.Ordinal];
                cells.FillPattern = row[ this.FillPattern.Ordinal].ToInt();
                cells.ShapeShdwObliqueAngle = row[ this.ShapeShdwObliqueAngle.Ordinal];
                cells.ShapeShdwOffsetX = row[ this.ShapeShdwOffsetX.Ordinal];
                cells.ShapeShdwOffsetY = row[ this.ShapeShdwOffsetY.Ordinal];
                cells.ShapeShdwScaleFactor = row[ this.ShapeShdwScaleFactor.Ordinal];
                cells.ShapeShdwType = row[ this.ShapeShdwType.Ordinal].ToInt();
                cells.ShdwBkgnd = row[ this.ShdwBkgnd.Ordinal].ToInt();
                cells.ShdwBkgndTrans = row[ this.ShdwBkgndTrans.Ordinal];
                cells.ShdwForegnd = row[ this.ShdwForegnd.Ordinal].ToInt();
                cells.ShdwForegndTrans = row[ this.ShdwForegndTrans.Ordinal];
                cells.ShdwPattern = row[ this.ShdwPattern.Ordinal].ToInt();
                cells.BeginArrow = row[ this.BeginArrow.Ordinal].ToInt();
                cells.BeginArrowSize = row[ this.BeginArrowSize.Ordinal];
                cells.EndArrow = row[ this.EndArrow.Ordinal].ToInt();
                cells.EndArrowSize = row[ this.EndArrowSize.Ordinal];
                cells.LineCap = row[ this.LineCap.Ordinal].ToInt();
                cells.LineColor = row[ this.LineColor.Ordinal].ToInt();
                cells.LineColorTrans = row[ this.LineColorTrans.Ordinal];
                cells.LinePattern = row[ this.LinePattern.Ordinal].ToInt();
                cells.LineWeight = row[ this.LineWeight.Ordinal];
                cells.Rounding = row[ this.Rounding.Ordinal];
                return cells;
            }

        }
    }
}