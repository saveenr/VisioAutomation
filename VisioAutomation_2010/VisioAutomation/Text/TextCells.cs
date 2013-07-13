using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Text
{
    public class TextCells : VA.ShapeSheet.CellGroups.CellGroupEx
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
            return VA.ShapeSheet.CellGroups.CellGroupEx._GetCells(page, shapeids, query, query.GetCells);
        }

        public static TextCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroupEx._GetCells(shape, query, query.GetCells);
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
            public int BottomMargin { get; set; }
            public int LeftMargin { get; set; }
            public int RightMargin { get; set; }
            public int TopMargin { get; set; }
            public int DefaultTabStop { get; set; }
            public int TextBkgnd { get; set; }
            public int TextBkgndTrans { get; set; }
            public int TextDirection { get; set; }
            public int VerticalAlign { get; set; }
            public int TxtWidth { get; set; }
            public int TxtHeight { get; set; }
            public int TxtPinX { get; set; }
            public int TxtPinY { get; set; }
            public int TxtLocPinX { get; set; }
            public int TxtLocPinY { get; set; }
            public int TxtAngle { get; set; }

            public TextBlockFormatQuery() :
                base()
            {
                BottomMargin = this.AddCell(VA.ShapeSheet.SRCConstants.BottomMargin, "BottomMargin");
                LeftMargin = this.AddCell(VA.ShapeSheet.SRCConstants.LeftMargin, "LeftMargin");
                RightMargin = this.AddCell(VA.ShapeSheet.SRCConstants.RightMargin, "RightMargin");
                TopMargin = this.AddCell(VA.ShapeSheet.SRCConstants.TopMargin, "TopMargin");
                DefaultTabStop = this.AddCell(VA.ShapeSheet.SRCConstants.DefaultTabStop, "DefaultTabStop");
                TextBkgnd = this.AddCell(VA.ShapeSheet.SRCConstants.TextBkgnd, "TextBkgnd");
                TextBkgndTrans = this.AddCell(VA.ShapeSheet.SRCConstants.TextBkgndTrans, "TextBkgndTrans");
                TextDirection = this.AddCell(VA.ShapeSheet.SRCConstants.TextDirection, "TextDirection");
                VerticalAlign = this.AddCell(VA.ShapeSheet.SRCConstants.VerticalAlign, "VerticalAlign");
                TxtPinX = this.AddCell(VA.ShapeSheet.SRCConstants.TxtPinX, "TxtPinX");
                TxtPinY = this.AddCell(VA.ShapeSheet.SRCConstants.TxtPinY, "TxtPinY");
                TxtLocPinX = this.AddCell(VA.ShapeSheet.SRCConstants.TxtLocPinX, "TxtLocPinX");
                TxtLocPinY = this.AddCell(VA.ShapeSheet.SRCConstants.TxtLocPinY, "TxtLocPinY");
                TxtWidth = this.AddCell(VA.ShapeSheet.SRCConstants.TxtWidth, "TxtWidth");
                TxtHeight = this.AddCell(VA.ShapeSheet.SRCConstants.TxtHeight, "TxtHeight");
                TxtAngle = this.AddCell(VA.ShapeSheet.SRCConstants.TxtAngle, "TxtAngle");
            }

            public TextCells GetCells(ExQueryResult<CellData<double>> data_for_shape)
            {
                var row = data_for_shape.Cells;

                var cells = new TextCells();
                cells.BottomMargin = row[BottomMargin];
                cells.LeftMargin = row[LeftMargin];
                cells.RightMargin = row[RightMargin];
                cells.TopMargin = row[TopMargin];
                cells.DefaultTabStop = row[DefaultTabStop];
                cells.TextBkgnd = row[TextBkgnd].ToInt();
                cells.TextBkgndTrans = row[TextBkgndTrans];
                cells.TextDirection = row[TextDirection].ToInt();
                cells.VerticalAlign = row[VerticalAlign].ToInt();
                cells.TxtPinX = row[TxtPinX];
                cells.TxtPinY = row[TxtPinY];
                cells.TxtLocPinX = row[TxtLocPinX];
                cells.TxtLocPinY = row[TxtLocPinY];
                cells.TxtWidth = row[TxtWidth];
                cells.TxtHeight = row[TxtHeight];
                cells.TxtAngle = row[TxtAngle];
                return cells;
            }
        }
    }
}