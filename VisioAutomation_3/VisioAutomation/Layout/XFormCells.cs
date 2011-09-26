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

        private static XFormCells get_cells_from_row(XFormQuery query, VA.ShapeSheet.Query.QueryDataRow<double> row)
        {
            var cells = new XFormCells();
            cells.PinX = row.GetItem(query.PinX);
            cells.PinY = row.GetItem(query.PinY);
            cells.LocPinX = row.GetItem(query.LocPinX);
            cells.LocPinY = row.GetItem(query.LocPinY);
            cells.Width = row.GetItem(query.Width);
            cells.Height = row.GetItem(query.Height);
            cells.Angle = row.GetItem(query.Angle);
            return cells;
        }

        internal static IList<XFormCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new XFormQuery();
            return VA.ShapeSheet.CellDataGroup._GetObjectsFromRows(page, shapeids, query, get_cells_from_row);
        }

        internal static XFormCells GetCells(IVisio.Shape shape)
        {
            var query = new XFormQuery();
            return VA.ShapeSheet.CellDataGroup._GetObjectFromSingleRow(shape, query, get_cells_from_row);
        }

        class XFormQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.CellQueryColumn Width { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn Height { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PinX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn PinY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LocPinX { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn LocPinY { get; set; }
            public VA.ShapeSheet.Query.CellQueryColumn Angle { get; set; }

            public XFormQuery() :
                base()
            {
                PinX = this.AddColumn(VA.ShapeSheet.SRCConstants.PinX, "PinX");
                PinY = this.AddColumn(VA.ShapeSheet.SRCConstants.PinY, "PinY");
                LocPinX = this.AddColumn(VA.ShapeSheet.SRCConstants.LocPinX, "LocPinX");
                LocPinY = this.AddColumn(VA.ShapeSheet.SRCConstants.LocPinY, "LocPinY");
                Width = this.AddColumn(VA.ShapeSheet.SRCConstants.Width, "Width");
                Height = this.AddColumn(VA.ShapeSheet.SRCConstants.Height, "Height");
                Angle = this.AddColumn(VA.ShapeSheet.SRCConstants.Angle, "Angle");
            }
        }

    }
}