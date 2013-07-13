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

        private static TextBlockFormatQuery m_query;
        private static TextBlockFormatQuery get_query()
        {
            m_query= m_query ?? new TextBlockFormatQuery();
            return m_query;
        }

        private static TextCells get_cells_from_row(TextBlockFormatQuery query,
                                                               VA.ShapeSheet.Data.Table <VA.ShapeSheet.CellData<double>> table, int row)
        {
            var cells = new TextCells();
            cells.BottomMargin = table[row,query.BottomMargin];
            cells.LeftMargin = table[row,query.LeftMargin];
            cells.RightMargin = table[row,query.RightMargin];
            cells.TopMargin = table[row,query.TopMargin];
            cells.DefaultTabStop = table[row,query.DefaultTabStop];
            cells.TextBkgnd = table[row,query.TextBkgnd].ToInt();
            cells.TextBkgndTrans = table[row,query.TextBkgndTrans];
            cells.TextDirection = table[row,query.TextDirection].ToInt();
            cells.VerticalAlign = table[row,query.VerticalAlign].ToInt();
            cells.TxtPinX = table[row,query.TxtPinX];
            cells.TxtPinY = table[row,query.TxtPinY];
            cells.TxtLocPinX = table[row,query.TxtLocPinX];
            cells.TxtLocPinY = table[row,query.TxtLocPinY];
            cells.TxtWidth = table[row,query.TxtWidth];
            cells.TxtHeight = table[row,query.TxtHeight];
            cells.TxtAngle = table[row,query.TxtAngle];
            return cells;
        }

        private class TextBlockFormatQuery : VA.ShapeSheet.Query.QueryEx
        {
            public QueryColumn BottomMargin { get; set; }
            public QueryColumn LeftMargin { get; set; }
            public QueryColumn RightMargin { get; set; }
            public QueryColumn TopMargin { get; set; }
            public QueryColumn DefaultTabStop { get; set; }
            public QueryColumn TextBkgnd { get; set; }
            public QueryColumn TextBkgndTrans { get; set; }
            public QueryColumn TextDirection { get; set; }
            public QueryColumn VerticalAlign { get; set; }
            public QueryColumn TxtWidth { get; set; }
            public QueryColumn TxtHeight { get; set; }
            public QueryColumn TxtPinX { get; set; }
            public QueryColumn TxtPinY { get; set; }
            public QueryColumn TxtLocPinX { get; set; }
            public QueryColumn TxtLocPinY { get; set; }
            public QueryColumn TxtAngle { get; set; }

            public TextBlockFormatQuery() :
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

            public TextCells GetCells(ExQueryResult<CellData<double>> data_for_shape)
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