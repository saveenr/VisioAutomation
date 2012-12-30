using System.Collections.Generic;
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

        protected override void ApplyFormulas(ApplyFormula func)
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
            return VA.ShapeSheet.CellGroups.CellGroup.CellsFromRows(page, shapeids, query, get_cells_from_row);
        }

        public static TextCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup.CellsFromRow(shape, query, get_cells_from_row);
        }

        private static TextBlockFormatQuery m_query;
        private static TextBlockFormatQuery get_query()
        {
            if (m_query==null)
            {
                m_query = new TextBlockFormatQuery();
            }
            return m_query;
        }

        private static TextCells get_cells_from_row(TextBlockFormatQuery query,
                                                               VA.ShapeSheet.Data.TableRow
                                                                   <VA.ShapeSheet.CellData<double>> row)
        {
            var cells = new TextCells();
            cells.BottomMargin = row[query.BottomMargin];
            cells.LeftMargin = row[query.LeftMargin];
            cells.RightMargin = row[query.RightMargin];
            cells.TopMargin = row[query.TopMargin];
            cells.DefaultTabStop = row[query.DefaultTabStop];
            cells.TextBkgnd = row[query.TextBkgnd].ToInt();
            cells.TextBkgndTrans = row[query.TextBkgndTrans];
            cells.TextDirection = row[query.TextDirection].ToInt();
            cells.VerticalAlign = row[query.VerticalAlign].ToInt();
            cells.TxtPinX = row[query.TxtPinX];
            cells.TxtPinY = row[query.TxtPinY];
            cells.TxtLocPinX = row[query.TxtLocPinX];
            cells.TxtLocPinY = row[query.TxtLocPinY];
            cells.TxtWidth = row[query.TxtWidth];
            cells.TxtHeight = row[query.TxtHeight];
            cells.TxtAngle = row[query.TxtAngle];
            return cells;
        }

        private class TextBlockFormatQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.QueryColumn BottomMargin { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LeftMargin { get; set; }
            public VA.ShapeSheet.Query.QueryColumn RightMargin { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TopMargin { get; set; }
            public VA.ShapeSheet.Query.QueryColumn DefaultTabStop { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TextBkgnd { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TextBkgndTrans { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TextDirection { get; set; }
            public VA.ShapeSheet.Query.QueryColumn VerticalAlign { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtWidth { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtHeight { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtPinX { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtPinY { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtLocPinX { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtLocPinY { get; set; }
            public VA.ShapeSheet.Query.QueryColumn TxtAngle { get; set; }

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
        }
    }
}