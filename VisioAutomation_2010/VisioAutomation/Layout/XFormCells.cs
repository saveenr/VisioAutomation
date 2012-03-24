using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{
    public partial class XFormCells : VA.ShapeSheet.CellGroups.CellGroup
    {
        public VA.ShapeSheet.CellData<double> PinX { get; set; }
        public VA.ShapeSheet.CellData<double> PinY { get; set; }
        public VA.ShapeSheet.CellData<double> LocPinX { get; set; }
        public VA.ShapeSheet.CellData<double> LocPinY { get; set; }
        public VA.ShapeSheet.CellData<double> Width { get; set; }
        public VA.ShapeSheet.CellData<double> Height { get; set; }
        public VA.ShapeSheet.CellData<double> Angle { get; set; }

        protected override void ApplyFormulas(ApplyFormula func)
        {
            func(ShapeSheet.SRCConstants.PinX, this.PinX.Formula);
            func(ShapeSheet.SRCConstants.PinY, this.PinY.Formula);
            func(ShapeSheet.SRCConstants.LocPinX, this.LocPinX.Formula);
            func(ShapeSheet.SRCConstants.LocPinY, this.LocPinY.Formula);
            func(ShapeSheet.SRCConstants.Width, this.Width.Formula);
            func(ShapeSheet.SRCConstants.Height, this.Height.Formula);
            func(ShapeSheet.SRCConstants.Angle, this.Angle.Formula);
        }

        private static XFormCells get_cells_from_row(XFormQuery query, VA.ShapeSheet.Data.TableRow<VA.ShapeSheet.CellData<double>> row)
        {
            var cells = new XFormCells();
            cells.PinX = row[query.PinX];
            cells.PinY = row[query.PinY];
            cells.LocPinX = row[query.LocPinX];
            cells.LocPinY = row[query.LocPinY];
            cells.Width = row[query.Width];
            cells.Height = row[query.Height];
            cells.Angle = row[query.Angle];
            return cells;
        }

        internal static IList<XFormCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new XFormQuery();
            return VA.ShapeSheet.CellGroups.CellGroup.CellsFromRows(page, shapeids, query, get_cells_from_row);
        }

        internal static XFormCells GetCells(IVisio.Shape shape)
        {
            var query = new XFormQuery();
            return VA.ShapeSheet.CellGroups.CellGroup.CellsFromRow(shape, query, get_cells_from_row);
        }

        class XFormQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.QueryColumn Width { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Height { get; set; }
            public VA.ShapeSheet.Query.QueryColumn PinX { get; set; }
            public VA.ShapeSheet.Query.QueryColumn PinY { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LocPinX { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LocPinY { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Angle { get; set; }

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