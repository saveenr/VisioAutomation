using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Text
{
    public class TextCells : VA.ShapeSheet.CellGroups.CellGroup
    {
        public VA.ShapeSheet.CellData<double> BottomMargin { get; set; }
        public VA.ShapeSheet.CellData<double> LeftMargin { get; set; }
        public VA.ShapeSheet.CellData<double> RightMargin { get; set; }
        public VA.ShapeSheet.CellData<double> TopMargin { get; set; }
        public VA.ShapeSheet.CellData<double> DefaultTabStop { get; set; }
        public VA.ShapeSheet.CellData<int> TextBkgnd { get; set; }
        public VA.ShapeSheet.CellData<double> TextBkgndTrans { get; set; }
        public VA.ShapeSheet.CellData<int> TextDirection { get; set; }
        public VA.ShapeSheet.CellData<int> VerticalAlign { get; set; }
        public VA.ShapeSheet.CellData<double> TxtAngle { get; set; }
        public VA.ShapeSheet.CellData<double> TxtWidth { get; set; }
        public VA.ShapeSheet.CellData<double> TxtHeight { get; set; }
        public VA.ShapeSheet.CellData<double> TxtPinX { get; set; }
        public VA.ShapeSheet.CellData<double> TxtPinY { get; set; }
        public VA.ShapeSheet.CellData<double> TxtLocPinX { get; set; }
        public VA.ShapeSheet.CellData<double> TxtLocPinY { get; set; }

        public override void ApplyFormulas(ApplyFormula func)
        {
            func(VA.ShapeSheet.SRCConstants.BottomMargin, this.BottomMargin.Formula);
            func(VA.ShapeSheet.SRCConstants.LeftMargin, this.LeftMargin.Formula);
            func(VA.ShapeSheet.SRCConstants.RightMargin, this.RightMargin.Formula);
            func(VA.ShapeSheet.SRCConstants.TopMargin, this.TopMargin.Formula);
            func(VA.ShapeSheet.SRCConstants.DefaultTabStop, this.DefaultTabStop.Formula);
            func(VA.ShapeSheet.SRCConstants.TextBkgnd, this.TextBkgnd.Formula);
            func(VA.ShapeSheet.SRCConstants.TextBkgndTrans, this.TextBkgndTrans.Formula);
            func(VA.ShapeSheet.SRCConstants.TextDirection, this.TextDirection.Formula);
            func(VA.ShapeSheet.SRCConstants.VerticalAlign, this.VerticalAlign.Formula);
            func(VA.ShapeSheet.SRCConstants.TxtPinX, this.TxtPinX.Formula);
            func(VA.ShapeSheet.SRCConstants.TxtPinY, this.TxtPinY.Formula);
            func(VA.ShapeSheet.SRCConstants.TxtLocPinX, this.TxtLocPinX.Formula);
            func(VA.ShapeSheet.SRCConstants.TxtLocPinY, this.TxtLocPinY.Formula);
            func(VA.ShapeSheet.SRCConstants.TxtWidth, this.TxtWidth.Formula);
            func(VA.ShapeSheet.SRCConstants.TxtHeight, this.TxtHeight.Formula);
            func(VA.ShapeSheet.SRCConstants.TxtAngle, this.TxtAngle.Formula);
        }

        public static IList<TextCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells(page, shapeids, query, query.GetCells);
        }

        public static TextCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells(shape, query, query.GetCells);
        }

        private static TextBlockFormatCellQuery _mCellQuery;
        private static TextBlockFormatCellQuery get_query()
        {
            _mCellQuery= _mCellQuery ?? new TextBlockFormatCellQuery();
            return _mCellQuery;
        }

        private static TextCells get_cells_from_row(TextBlockFormatCellQuery cellQuery,
                                                               VA.ShapeSheet.Data.Table <VA.ShapeSheet.CellData<double>> table, int row)
        {
            var cells = new TextCells();
            cells.BottomMargin = table[row,cellQuery.BottomMargin];
            cells.LeftMargin = table[row,cellQuery.LeftMargin];
            cells.RightMargin = table[row,cellQuery.RightMargin];
            cells.TopMargin = table[row,cellQuery.TopMargin];
            cells.DefaultTabStop = table[row,cellQuery.DefaultTabStop];
            cells.TextBkgnd = table[row,cellQuery.TextBkgnd].ToInt();
            cells.TextBkgndTrans = table[row,cellQuery.TextBkgndTrans];
            cells.TextDirection = table[row,cellQuery.TextDirection].ToInt();
            cells.VerticalAlign = table[row,cellQuery.VerticalAlign].ToInt();
            cells.TxtPinX = table[row,cellQuery.TxtPinX];
            cells.TxtPinY = table[row,cellQuery.TxtPinY];
            cells.TxtLocPinX = table[row,cellQuery.TxtLocPinX];
            cells.TxtLocPinY = table[row,cellQuery.TxtLocPinY];
            cells.TxtWidth = table[row,cellQuery.TxtWidth];
            cells.TxtHeight = table[row,cellQuery.TxtHeight];
            cells.TxtAngle = table[row,cellQuery.TxtAngle];
            return cells;
        }

        private class TextBlockFormatCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public Column BottomMargin { get; set; }
            public Column LeftMargin { get; set; }
            public Column RightMargin { get; set; }
            public Column TopMargin { get; set; }
            public Column DefaultTabStop { get; set; }
            public Column TextBkgnd { get; set; }
            public Column TextBkgndTrans { get; set; }
            public Column TextDirection { get; set; }
            public Column VerticalAlign { get; set; }
            public Column TxtWidth { get; set; }
            public Column TxtHeight { get; set; }
            public Column TxtPinX { get; set; }
            public Column TxtPinY { get; set; }
            public Column TxtLocPinX { get; set; }
            public Column TxtLocPinY { get; set; }
            public Column TxtAngle { get; set; }

            public TextBlockFormatCellQuery() :
                base()
            {
                BottomMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.BottomMargin, "BottomMargin");
                LeftMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.LeftMargin, "LeftMargin");
                RightMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.RightMargin, "RightMargin");
                TopMargin = this.AddColumn(VA.ShapeSheet.SRCConstants.TopMargin, "TopMargin");
                DefaultTabStop = this.AddColumn(VA.ShapeSheet.SRCConstants.DefaultTabStop, "DefaultTabStop");
                TextBkgnd = this.AddColumn(VA.ShapeSheet.SRCConstants.TextBkgnd, "TextBkgnd");
                TextBkgndTrans = this.AddColumn(VA.ShapeSheet.SRCConstants.TextBkgndTrans, "TextBkgndTrans");
                TextDirection = this.AddColumn(VA.ShapeSheet.SRCConstants.TextDirection, "TextDirection");
                VerticalAlign = this.AddColumn(VA.ShapeSheet.SRCConstants.VerticalAlign, "VerticalAlign");
                TxtPinX = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtPinX, "TxtPinX");
                TxtPinY = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtPinY, "TxtPinY");
                TxtLocPinX = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtLocPinX, "TxtLocPinX");
                TxtLocPinY = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtLocPinY, "TxtLocPinY");
                TxtWidth = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtWidth, "TxtWidth");
                TxtHeight = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtHeight, "TxtHeight");
                TxtAngle = this.AddColumn(VA.ShapeSheet.SRCConstants.TxtAngle, "TxtAngle");
            }

            public TextCells GetCells(QueryResult<CellData<double>> data_for_shape)
            {
                var row = data_for_shape.Cells;

                var cells = new TextCells();
                cells.BottomMargin = row[BottomMargin.Ordinal];
                cells.LeftMargin = row[LeftMargin.Ordinal];
                cells.RightMargin = row[RightMargin.Ordinal];
                cells.TopMargin = row[TopMargin.Ordinal];
                cells.DefaultTabStop = row[DefaultTabStop.Ordinal];
                cells.TextBkgnd = row[TextBkgnd.Ordinal].ToInt();
                cells.TextBkgndTrans = row[TextBkgndTrans.Ordinal];
                cells.TextDirection = row[TextDirection.Ordinal].ToInt();
                cells.VerticalAlign = row[VerticalAlign.Ordinal].ToInt();
                cells.TxtPinX = row[TxtPinX.Ordinal];
                cells.TxtPinY = row[TxtPinY.Ordinal];
                cells.TxtLocPinX = row[TxtLocPinX.Ordinal];
                cells.TxtLocPinY = row[TxtLocPinY.Ordinal];
                cells.TxtWidth = row[TxtWidth.Ordinal];
                cells.TxtHeight = row[TxtHeight.Ordinal];
                cells.TxtAngle = row[TxtAngle.Ordinal];
                return cells;
            }
        }
    }
}