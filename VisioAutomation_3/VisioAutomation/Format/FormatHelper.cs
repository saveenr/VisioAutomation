using VA=VisioAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Format
{
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