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

        public override IEnumerable<VA.ShapeSheet.CellGroups.BaseCellGroup.SRCValuePair> EnumPairs()
        {
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.BottomMargin, this.BottomMargin.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.LeftMargin, this.LeftMargin.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.RightMargin, this.RightMargin.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.TopMargin, this.TopMargin.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.DefaultTabStop, this.DefaultTabStop.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.TextBkgnd, this.TextBkgnd.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.TextBkgndTrans, this.TextBkgndTrans.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.TextDirection, this.TextDirection.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.VerticalAlign, this.VerticalAlign.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.TxtPinX, this.TxtPinX.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.TxtPinY, this.TxtPinY.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.TxtLocPinX, this.TxtLocPinX.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.TxtLocPinY, this.TxtLocPinY.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.TxtWidth, this.TxtWidth.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.TxtHeight, this.TxtHeight.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.TxtAngle, this.TxtAngle.Formula);
        }

        public static IList<TextCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells<TextCells, double>(page, shapeids, query, query.GetCells);
        }

        public static TextCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells<TextCells, double>(shape, query, query.GetCells);
        }

        private static TextBlockFormatCellQuery _mCellQuery;
        private static TextBlockFormatCellQuery get_query()
        {
            _mCellQuery= _mCellQuery ?? new TextBlockFormatCellQuery();
            return _mCellQuery;
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
                BottomMargin = this.Columns.Add(VA.ShapeSheet.SRCConstants.BottomMargin, "BottomMargin");
                LeftMargin = this.Columns.Add(VA.ShapeSheet.SRCConstants.LeftMargin, "LeftMargin");
                RightMargin = this.Columns.Add(VA.ShapeSheet.SRCConstants.RightMargin, "RightMargin");
                TopMargin = this.Columns.Add(VA.ShapeSheet.SRCConstants.TopMargin, "TopMargin");
                DefaultTabStop = this.Columns.Add(VA.ShapeSheet.SRCConstants.DefaultTabStop, "DefaultTabStop");
                TextBkgnd = this.Columns.Add(VA.ShapeSheet.SRCConstants.TextBkgnd, "TextBkgnd");
                TextBkgndTrans = this.Columns.Add(VA.ShapeSheet.SRCConstants.TextBkgndTrans, "TextBkgndTrans");
                TextDirection = this.Columns.Add(VA.ShapeSheet.SRCConstants.TextDirection, "TextDirection");
                VerticalAlign = this.Columns.Add(VA.ShapeSheet.SRCConstants.VerticalAlign, "VerticalAlign");
                TxtPinX = this.Columns.Add(VA.ShapeSheet.SRCConstants.TxtPinX, "TxtPinX");
                TxtPinY = this.Columns.Add(VA.ShapeSheet.SRCConstants.TxtPinY, "TxtPinY");
                TxtLocPinX = this.Columns.Add(VA.ShapeSheet.SRCConstants.TxtLocPinX, "TxtLocPinX");
                TxtLocPinY = this.Columns.Add(VA.ShapeSheet.SRCConstants.TxtLocPinY, "TxtLocPinY");
                TxtWidth = this.Columns.Add(VA.ShapeSheet.SRCConstants.TxtWidth, "TxtWidth");
                TxtHeight = this.Columns.Add(VA.ShapeSheet.SRCConstants.TxtHeight, "TxtHeight");
                TxtAngle = this.Columns.Add(VA.ShapeSheet.SRCConstants.TxtAngle, "TxtAngle");
            }

            public TextCells GetCells(VA.ShapeSheet.CellData<double>[] row)
            {
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