using VA=VisioAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Format
{
    public class ShapeFormatCells : VA.ShapeSheet.CellDataGroup
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

        protected override void _Apply(VA.ShapeSheet.CellDataGroup.ApplyFormula func)
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

        private static ShapeFormatCells get_cells_from_row(ShapeFormatQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new ShapeFormatCells();
            cells.FillBkgnd = qds.GetItem(row, query.FillBkgnd, v => (int)v);
            cells.FillBkgndTrans = qds.GetItem(row, query.FillBkgndTrans);
            cells.FillForegnd = qds.GetItem(row, query.FillForegnd, v => (int)v);
            cells.FillForegndTrans = qds.GetItem(row, query.FillForegndTrans);
            cells.FillPattern = qds.GetItem(row, query.FillPattern, v => (int)v);
            cells.ShapeShdwObliqueAngle = qds.GetItem(row, query.ShapeShdwObliqueAngle);
            cells.ShapeShdwOffsetX = qds.GetItem(row, query.ShapeShdwOffsetX);
            cells.ShapeShdwOffsetY = qds.GetItem(row, query.ShapeShdwOffsetY);
            cells.ShapeShdwScaleFactor = qds.GetItem(row, query.ShapeShdwScaleFactor);
            cells.ShapeShdwType = qds.GetItem(row, query.ShapeShdwType, v => (int)v);
            cells.ShdwBkgnd = qds.GetItem(row, query.ShdwBkgnd, v => (int)v);
            cells.ShdwBkgndTrans = qds.GetItem(row, query.ShdwBkgndTrans);
            cells.ShdwForegnd = qds.GetItem(row, query.ShdwForegnd, v => (int)v);
            cells.ShdwForegndTrans = qds.GetItem(row, query.ShdwForegndTrans);
            cells.ShdwPattern = qds.GetItem(row, query.ShdwPattern, v => (int)v);
            cells.BeginArrow = qds.GetItem(row, query.BeginArrow, v => (int)v);
            cells.BeginArrowSize = qds.GetItem(row, query.BeginArrowSize);
            cells.EndArrow = qds.GetItem(row, query.EndArrow, v => (int)v);
            cells.EndArrowSize = qds.GetItem(row, query.EndArrowSize);
            cells.LineCap = qds.GetItem(row, query.LineCap, v => (int)v);
            cells.LineColor = qds.GetItem(row, query.LineColor, v => (int)v);
            cells.LineColorTrans = qds.GetItem(row, query.LineColorTrans);
            cells.LinePattern = qds.GetItem(row, query.LinePattern, v => (int)v);
            cells.LineWeight = qds.GetItem(row, query.LineWeight);
            cells.Rounding = qds.GetItem(row, query.Rounding);
            cells.CharFont = qds.GetItem(row, query.CharFont, v => (int)v);
            cells.CharColor = qds.GetItem(row, query.CharColor, v => (int)v);
            cells.CharColorTrans = qds.GetItem(row, query.CharColorTrans);
            cells.CharSize = qds.GetItem(row, query.CharSize);
            cells.TextBkgnd = qds.GetItem(row, query.TextBkgnd, v => (int)v);
            cells.TextBkgndTrans = qds.GetItem(row, query.TextBkgndTrans);
            return cells;
        }

        internal static IList<ShapeFormatCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new ShapeFormatQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(page, shapeids, query, get_cells_from_row);
        }

        internal static ShapeFormatCells GetCells(IVisio.Shape shape)
        {
            var query = new ShapeFormatQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(shape, query, get_cells_from_row);
        }
    }
}