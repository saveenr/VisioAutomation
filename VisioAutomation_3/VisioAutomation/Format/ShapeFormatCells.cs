using VA=VisioAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Format
{
    public class ShapeFormatCells
    {
        // Fill
        public VA.ShapeSheet.CellData<int> FillBkgnd { get; set; }
        public VA.ShapeSheet.CellData<double>FillBkgndTrans { get; set; }
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

        // Line
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

        // Char
        public VA.ShapeSheet.CellData<int> CharFont { get; set; }
        public VA.ShapeSheet.CellData<int> CharColor { get; set; }
        public VA.ShapeSheet.CellData<double> CharColorTrans { get; set; }
        public VA.ShapeSheet.CellData<double> CharSize { get; set; }

        // Text

        public VA.ShapeSheet.CellData<int> TextBkgnd { get; set; }
        public VA.ShapeSheet.CellData<double> TextBkgndTrans { get; set; }

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short id)
        {

            // Fill
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillBkgnd, FillBkgnd.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillBkgndTrans, FillBkgndTrans.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillForegnd, FillForegnd.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillForegndTrans, FillForegndTrans.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.FillPattern, FillPattern.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, ShapeShdwObliqueAngle.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwOffsetX, ShapeShdwOffsetX.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwOffsetY, ShapeShdwOffsetY.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwScaleFactor, ShapeShdwScaleFactor.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShapeShdwType, ShapeShdwType.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwBkgnd, ShdwBkgnd.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwBkgndTrans, ShdwBkgndTrans.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwForegnd, ShdwForegnd.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwForegndTrans, ShdwForegndTrans.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ShdwPattern, ShdwPattern.Formula);


            // Line
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BeginArrow, BeginArrow.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BeginArrowSize, BeginArrowSize.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineColor, LineColor.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineColorTrans, LineColorTrans.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LinePattern, LinePattern.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineWeight, LineWeight.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.EndArrow, EndArrow.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.EndArrowSize, EndArrowSize.Formula);

            // Char
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Char_Color, CharColor.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Char_ColorTrans, CharColorTrans.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Char_Font, CharFont.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.Char_Size, CharSize.Formula);

            // Text
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TextBkgnd, TextBkgnd.Formula);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.TextBkgndTrans, TextBkgndTrans.Formula);
        }
    }

    public static class FormatHelper
    {
        public static VA.Format.ShapeFormatCells GetShapeFormat(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var query = new ShapeFormatQuery();
            var qds = query.GetFormulasAndResults<double>(shape);
            var data = get_formatadata_for_row(query, qds, 0);

            return data;
        }

        public static IList<VA.Format.ShapeFormatCells> GetShapeFormat(IVisio.Page page, IList<int> shapeids)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var query = new ShapeFormatQuery();
            var qds = query.GetFormulasAndResults<double>(page, shapeids);
            var formats = new List<VA.Format.ShapeFormatCells>(shapeids.Count);
            for (int i = 0; i < qds.RowCount; i++)
            {
                var data = get_formatadata_for_row(query, qds, i);
                formats.Add(data);
            }

            return formats;
        }

        private static VA.Format.ShapeFormatCells get_formatadata_for_row(ShapeFormatQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var data = new VA.Format.ShapeFormatCells();

            data.FillBkgnd = qds.GetItem(row, query.FillBkgnd, v => (int)v);
            data.FillBkgndTrans = qds.GetItem(row, query.FillBkgndTrans);
            data.FillForegnd = qds.GetItem(row, query.FillForegnd, v => (int)v);
            data.FillForegndTrans = qds.GetItem(row, query.FillForegndTrans);
            data.FillPattern = qds.GetItem(row, query.FillPattern, v => (int)v);
            data.ShapeShdwObliqueAngle = qds.GetItem(row, query.ShapeShdwObliqueAngle);
            data.ShapeShdwOffsetX = qds.GetItem(row, query.ShapeShdwOffsetX);
            data.ShapeShdwOffsetY = qds.GetItem(row, query.ShapeShdwOffsetY);
            data.ShapeShdwScaleFactor = qds.GetItem(row, query.ShapeShdwScaleFactor);
            data.ShapeShdwType = qds.GetItem(row, query.ShapeShdwType, v => (int)v);
            data.ShdwBkgnd = qds.GetItem(row, query.ShdwBkgnd, v => (int)v);
            data.ShdwBkgndTrans = qds.GetItem(row, query.ShdwBkgndTrans);
            data.ShdwForegnd = qds.GetItem(row, query.ShdwForegnd, v => (int)v);
            data.ShdwForegndTrans = qds.GetItem(row, query.ShdwForegndTrans);
            data.ShdwPattern = qds.GetItem(row, query.ShdwPattern, v => (int)v);

            data.BeginArrow = qds.GetItem(row, query.BeginArrow, v => (int)v);
            data.BeginArrowSize = qds.GetItem(row, query.BeginArrowSize);
            data.EndArrow = qds.GetItem(row, query.EndArrow, v => (int)v);
            data.EndArrowSize = qds.GetItem(row, query.EndArrowSize);

            data.LineCap = qds.GetItem(row, query.LineCap, v => (int)v);
            data.LineColor = qds.GetItem(row, query.LineColor, v => (int)v);
            data.LinePattern = qds.GetItem(row, query.LinePattern, v => (int)v);
            data.LineColorTrans = qds.GetItem(row, query.LineColorTrans);
            data.LineWeight = qds.GetItem(row, query.LineWeight);

            data.CharColor = qds.GetItem(row, query.CharColor, v => (int)v);
            data.CharColorTrans = qds.GetItem(row, query.CharColorTrans);
            data.CharFont = qds.GetItem(row, query.CharFont, v => (int)v);
            data.CharSize = qds.GetItem(row, query.CharSize);

            data.TextBkgnd = qds.GetItem(row, query.TextBkgnd, v => (int)v);
            data.TextBkgndTrans = qds.GetItem(row, query.TextBkgndTrans);

            return data;
        }
        
    }
}