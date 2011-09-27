using VA=VisioAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Format
{
    public class ShapeFormatCells : VA.ShapeSheet.CellGroups.CellGroup
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
        public VA.ShapeSheet.CellData<int> CharFont { get; set; }
        public VA.ShapeSheet.CellData<int> CharColor { get; set; }
        public VA.ShapeSheet.CellData<double> CharColorTrans { get; set; }
        public VA.ShapeSheet.CellData<double> CharSize { get; set; }
        public VA.ShapeSheet.CellData<int> TextBkgnd { get; set; }
        public VA.ShapeSheet.CellData<double> TextBkgndTrans { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellGroups.CellGroup.ApplyFormula func)
        {
            func(ShapeSheet.SRCConstants.FillBkgnd, this.FillBkgnd.Formula);
            func(ShapeSheet.SRCConstants.FillBkgndTrans, this.FillBkgndTrans.Formula);
            func(ShapeSheet.SRCConstants.FillForegnd, this.FillForegnd.Formula);
            func(ShapeSheet.SRCConstants.FillForegndTrans, this.FillForegndTrans.Formula);
            func(ShapeSheet.SRCConstants.FillPattern, this.FillPattern.Formula);
            func(ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, this.ShapeShdwObliqueAngle.Formula);
            func(ShapeSheet.SRCConstants.ShapeShdwOffsetX, this.ShapeShdwOffsetX.Formula);
            func(ShapeSheet.SRCConstants.ShapeShdwOffsetY, this.ShapeShdwOffsetY.Formula);
            func(ShapeSheet.SRCConstants.ShapeShdwScaleFactor, this.ShapeShdwScaleFactor.Formula);
            func(ShapeSheet.SRCConstants.ShapeShdwType, this.ShapeShdwType.Formula);
            func(ShapeSheet.SRCConstants.ShdwBkgnd, this.ShdwBkgnd.Formula);
            func(ShapeSheet.SRCConstants.ShdwBkgndTrans, this.ShdwBkgndTrans.Formula);
            func(ShapeSheet.SRCConstants.ShdwForegnd, this.ShdwForegnd.Formula);
            func(ShapeSheet.SRCConstants.ShdwForegndTrans, this.ShdwForegndTrans.Formula);
            func(ShapeSheet.SRCConstants.ShdwPattern, this.ShdwPattern.Formula);
            func(ShapeSheet.SRCConstants.BeginArrow, this.BeginArrow.Formula);
            func(ShapeSheet.SRCConstants.BeginArrowSize, this.BeginArrowSize.Formula);
            func(ShapeSheet.SRCConstants.EndArrow, this.EndArrow.Formula);
            func(ShapeSheet.SRCConstants.EndArrowSize, this.EndArrowSize.Formula);
            func(ShapeSheet.SRCConstants.LineCap, this.LineCap.Formula);
            func(ShapeSheet.SRCConstants.LineColor, this.LineColor.Formula);
            func(ShapeSheet.SRCConstants.LineColorTrans, this.LineColorTrans.Formula);
            func(ShapeSheet.SRCConstants.LinePattern, this.LinePattern.Formula);
            func(ShapeSheet.SRCConstants.LineWeight, this.LineWeight.Formula);
            func(ShapeSheet.SRCConstants.Rounding, this.Rounding.Formula);
            func(ShapeSheet.SRCConstants.Char_Font, this.CharFont.Formula);
            func(ShapeSheet.SRCConstants.Char_Color, this.CharColor.Formula);
            func(ShapeSheet.SRCConstants.Char_ColorTrans, this.CharColorTrans.Formula);
            func(ShapeSheet.SRCConstants.Char_Size, this.CharSize.Formula);
            func(ShapeSheet.SRCConstants.TextBkgnd, this.TextBkgnd.Formula);
            func(ShapeSheet.SRCConstants.TextBkgndTrans, this.TextBkgndTrans.Formula);
        }

        private static ShapeFormatCells get_cells_from_row(ShapeFormatQuery query, VA.ShapeSheet.Data.QueryDataRow<double> row)
        {

            var cells = new ShapeFormatCells();
            cells.FillBkgnd = row[query.FillBkgnd].ToInt();
            cells.FillBkgndTrans = row[query.FillBkgndTrans];
            cells.FillForegnd = row[query.FillForegnd].ToInt();
            cells.FillForegndTrans = row[query.FillForegndTrans];
            cells.FillPattern = row[query.FillPattern].ToInt();
            cells.ShapeShdwObliqueAngle = row[query.ShapeShdwObliqueAngle];
            cells.ShapeShdwOffsetX = row[query.ShapeShdwOffsetX];
            cells.ShapeShdwOffsetY = row[query.ShapeShdwOffsetY];
            cells.ShapeShdwScaleFactor = row[query.ShapeShdwScaleFactor];
            cells.ShapeShdwType = row[query.ShapeShdwType].ToInt();
            cells.ShdwBkgnd = row[query.ShdwBkgnd].ToInt();
            cells.ShdwBkgndTrans = row[query.ShdwBkgndTrans];
            cells.ShdwForegnd = row[query.ShdwForegnd].ToInt();
            cells.ShdwForegndTrans = row[query.ShdwForegndTrans];
            cells.ShdwPattern = row[query.ShdwPattern].ToInt();
            cells.BeginArrow = row[query.BeginArrow].ToInt();
            cells.BeginArrowSize = row[query.BeginArrowSize];
            cells.EndArrow = row[query.EndArrow].ToInt();
            cells.EndArrowSize = row[query.EndArrowSize];
            cells.LineCap = row[query.LineCap].ToInt();
            cells.LineColor = row[query.LineColor].ToInt();
            cells.LineColorTrans = row[query.LineColorTrans];
            cells.LinePattern = row[query.LinePattern].ToInt();
            cells.LineWeight = row[query.LineWeight];
            cells.Rounding = row[query.Rounding];
            cells.CharFont = row[query.CharFont].ToInt();
            cells.CharColor = row[query.CharColor].ToInt();
            cells.CharColorTrans = row[query.CharColorTrans];
            cells.CharSize = row[query.CharSize];
            cells.TextBkgnd = row[query.TextBkgnd].ToInt();
            cells.TextBkgndTrans = row[query.TextBkgndTrans];
            return cells;
        }

        internal static IList<ShapeFormatCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new ShapeFormatQuery();
            return VA.ShapeSheet.CellGroups.CellGroup._GetObjectsFromRows(page, shapeids, query, get_cells_from_row);
        }

        internal static ShapeFormatCells GetCells(IVisio.Shape shape)
        {
            var query = new ShapeFormatQuery();
            return VA.ShapeSheet.CellGroups.CellGroup._GetObjectFromSingleRow(shape, query, get_cells_from_row);
        }

        class ShapeFormatQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.CellQueryColumn FillBkgnd { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn FillBkgndTrans { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn FillForegnd { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn FillForegndTrans { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn FillPattern { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapeShdwObliqueAngle { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapeShdwOffsetX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapeShdwOffsetY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapeShdwScaleFactor { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShapeShdwType { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShdwBkgnd { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShdwBkgndTrans { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShdwForegnd { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShdwForegndTrans { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn ShdwPattern { get; set; }

            public VA.ShapeSheet.Query.CellQueryColumn BeginArrow { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn BeginArrowSize { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn EndArrow { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn EndArrowSize { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineColor { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineCap { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineColorTrans { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LinePattern { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LineWeight { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn Rounding { get; set; }

            public VA.ShapeSheet.Query.CellQueryColumn CharColor { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn CharColorTrans { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn CharSize { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn CharFont { get; set; }

            public VA.ShapeSheet.Query.CellQueryColumn TextBkgnd { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn TextBkgndTrans { get; set; }

            public ShapeFormatQuery() :
                base()
            {
                this.FillBkgnd = this.AddColumn(VA.ShapeSheet.SRCConstants.FillBkgnd, "FillBkgnd");
                this.FillBkgndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.FillBkgndTrans, "FillBkgndTrans");
                this.FillForegnd = this.AddColumn(VA.ShapeSheet.SRCConstants.FillForegnd, "FillForegnd");
                this.FillForegndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.FillForegndTrans, "FillForegndTrans");
                this.FillPattern = this.AddColumn(VA.ShapeSheet.SRCConstants.FillPattern, "FillPattern");
                this.ShapeShdwObliqueAngle = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, "ShapeShdwObliqueAngle");
                this.ShapeShdwOffsetX = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeShdwOffsetX, "ShapeShdwOffsetX");
                this.ShapeShdwOffsetY = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeShdwOffsetY, "ShapeShdwOffsetY");
                this.ShapeShdwScaleFactor = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeShdwScaleFactor, "ShapeShdwScaleFactor");
                this.ShapeShdwType = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeShdwType, "ShapeShdwType");
                this.ShdwBkgnd = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwBkgnd, "ShdwBkgnd ");
                this.ShdwBkgndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwBkgndTrans, "ShdwBkgndTrans");
                this.ShdwForegnd = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwForegnd, "ShdwForegnd ");
                this.ShdwForegndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwForegndTrans, "ShdwForegndTrans");
                this.ShdwPattern = this.AddColumn(VA.ShapeSheet.SRCConstants.ShdwPattern, "ShdwPattern");

                this.BeginArrow = this.AddColumn(VA.ShapeSheet.SRCConstants.BeginArrow, "BeginArrow");
                this.BeginArrowSize = this.AddColumn(VA.ShapeSheet.SRCConstants.BeginArrowSize, "BeginArrowSize");
                this.EndArrow = this.AddColumn(VA.ShapeSheet.SRCConstants.EndArrow, "EndArrow ");
                this.EndArrowSize = this.AddColumn(VA.ShapeSheet.SRCConstants.EndArrowSize, "EndArrowSize");
                this.LineColor = this.AddColumn(VA.ShapeSheet.SRCConstants.LineColor, "LineColor");
                this.LineCap = this.AddColumn(VA.ShapeSheet.SRCConstants.LineCap, "LineCap");
                this.LineColorTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.LineColorTrans, "LineColorTrans");
                this.LinePattern = this.AddColumn(VA.ShapeSheet.SRCConstants.LinePattern, "LinePattern");
                this.LineWeight = this.AddColumn(VA.ShapeSheet.SRCConstants.LineWeight, "LineWeight");
                this.Rounding = this.AddColumn(VA.ShapeSheet.SRCConstants.Rounding, "Rounding");

                this.CharColor = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Color, "CharColor");
                this.CharColorTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_ColorTrans, "CharColorTrans");
                this.CharSize = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Size, "CharSize");
                this.CharFont = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Font, "CharFont");

                this.TextBkgnd = this.AddColumn(VA.ShapeSheet.SRCConstants.TextBkgnd, "TextBkgnd");
                this.TextBkgndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.TextBkgndTrans, "TextBkgndTrans");
            }
        }

    }
}