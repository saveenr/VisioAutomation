using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{
    public partial class XFormCells : VA.ShapeSheet.CellDataGroup
    {
        public VA.ShapeSheet.CellData<double> PinX { get; set; }
        public VA.ShapeSheet.CellData<double> PinY { get; set; }
        public VA.ShapeSheet.CellData<double> LocPinX { get; set; }
        public VA.ShapeSheet.CellData<double> LocPinY { get; set; }
        public VA.ShapeSheet.CellData<double> Width { get; set; }
        public VA.ShapeSheet.CellData<double> Height { get; set; }
        public VA.ShapeSheet.CellData<double> Angle { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellDataGroup.ApplyFormula func)
        {

            func(ShapeSheet.SRCConstants.PinX, this.PinX.Formula);
            func(ShapeSheet.SRCConstants.PinY, this.PinY.Formula);
            func(ShapeSheet.SRCConstants.LocPinX, this.LocPinX.Formula);
            func(ShapeSheet.SRCConstants.LocPinY, this.LocPinY.Formula);
            func(ShapeSheet.SRCConstants.Width, this.Width.Formula);
            func(ShapeSheet.SRCConstants.Height, this.Height.Formula);
            func(ShapeSheet.SRCConstants.Angle, this.Angle.Formula);
        }

        private static XFormCells get_cells_from_row(XFormQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new XFormCells();
            cells.PinX = qds.GetItem(row, query.PinX);
            cells.PinY = qds.GetItem(row, query.PinY);
            cells.LocPinX = qds.GetItem(row, query.LocPinX);
            cells.LocPinY = qds.GetItem(row, query.LocPinY);
            cells.Width = qds.GetItem(row, query.Width);
            cells.Height = qds.GetItem(row, query.Height);
            cells.Angle = qds.GetItem(row, query.Angle);
            return cells;
        }

        public static IList<XFormCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new XFormQuery();
            var qds = query.GetFormulasAndResults<double>(page, shapeids);
            var cells_list = new List<XFormCells>(shapeids.Count);
            for (int i = 0; i < qds.RowCount; i++)
            {
                var cells = get_cells_from_row(query, qds, i);
                cells_list.Add(cells);
            }

            return cells_list;
        }

        public static XFormCells GetCells(IVisio.Shape shape)
        {
            var query = new XFormQuery();
            var qds = query.GetFormulasAndResults<double>(shape);
            var cells = get_cells_from_row(query, qds, 0);
            return cells;
        }

    }
}