using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Text
{
        public class TextXFormCells : VA.ShapeSheet.CellGroups.CellGroup
    {
        public VA.ShapeSheet.CellData<double> TxtAngle { get; set; }
        public VA.ShapeSheet.CellData<double> TxtWidth { get; set; }
        public VA.ShapeSheet.CellData<double> TxtHeight { get; set; }
        public VA.ShapeSheet.CellData<double> TxtPinX { get; set; }
        public VA.ShapeSheet.CellData<double> TxtPinY { get; set; }
        public VA.ShapeSheet.CellData<double> TxtLocPinX { get; set; }
        public VA.ShapeSheet.CellData<double> TxtLocPinY { get; set; }

        protected override void ApplyFormulas(ApplyFormula func)
        {
            func(ShapeSheet.SRCConstants.TxtPinX, this.TxtPinX.Formula);
            func(ShapeSheet.SRCConstants.TxtPinY, this.TxtPinY.Formula);
            func(ShapeSheet.SRCConstants.TxtLocPinX, this.TxtLocPinX.Formula);
            func(ShapeSheet.SRCConstants.TxtLocPinY, this.TxtLocPinY.Formula);
            func(ShapeSheet.SRCConstants.TxtWidth, this.TxtWidth.Formula);
            func(ShapeSheet.SRCConstants.TxtHeight, this.TxtHeight.Formula);
            func(ShapeSheet.SRCConstants.TxtAngle, this.TxtAngle.Formula);
        }

        private static TextXFormCells get_cells_from_row(TextXFormQuery query,
                                                         VA.ShapeSheet.Data.TableRow<VA.ShapeSheet.CellData<double>> row)
        {
            var cells = new TextXFormCells();
            cells.TxtPinX = row[query.TxtPinX];
            cells.TxtPinY = row[query.TxtPinY];
            cells.TxtLocPinX = row[query.TxtLocPinX];
            cells.TxtLocPinY = row[query.TxtLocPinY];
            cells.TxtWidth = row[query.TxtWidth];
            cells.TxtHeight = row[query.TxtHeight];
            cells.TxtAngle = row[query.TxtAngle];
            return cells;
        }

        internal static IList<TextXFormCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new TextXFormQuery();
            return VA.ShapeSheet.CellGroups.CellGroup.CellsFromRows(page, shapeids, query, get_cells_from_row);
        }

        internal static TextXFormCells GetCells(IVisio.Shape shape)
        {
            var query = new TextXFormQuery();
            return VA.ShapeSheet.CellGroups.CellGroup.CellsFromRow(shape, query, get_cells_from_row);
        }

        private class TextXFormQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.QueryColumn TxtWidth { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtHeight { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtPinX { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtPinY { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtLocPinX { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtLocPinY { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtAngle { get; set; }

            public TextXFormQuery() :
                base()
            {
                TxtPinX = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtPinX, "TxtPinX");
                TxtPinY = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtPinY, "TxtPinY");
                TxtLocPinX = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtLocPinX, "TxtLocPinX");
                TxtLocPinY = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtLocPinY, "TxtLocPinY");
                TxtWidth = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtWidth, "TxtWidth");
                TxtHeight = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtHeight, "TxtHeight");
                TxtAngle = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtAngle, "TxtAngle");
            }
        }
    }
}