using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

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

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return newpair(VA.ShapeSheet.SRCConstants.BottomMargin, this.BottomMargin.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.LeftMargin, this.LeftMargin.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.RightMargin, this.RightMargin.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.TopMargin, this.TopMargin.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.DefaultTabStop, this.DefaultTabStop.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.TextBkgnd, this.TextBkgnd.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.TextBkgndTrans, this.TextBkgndTrans.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.TextDirection, this.TextDirection.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.VerticalAlign, this.VerticalAlign.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.TxtPinX, this.TxtPinX.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.TxtPinY, this.TxtPinY.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.TxtLocPinX, this.TxtLocPinX.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.TxtLocPinY, this.TxtLocPinY.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.TxtWidth, this.TxtWidth.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.TxtHeight, this.TxtHeight.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.TxtAngle, this.TxtAngle.Formula);
            }
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
            public CellColumn BottomMargin { get; set; }
            public CellColumn LeftMargin { get; set; }
            public CellColumn RightMargin { get; set; }
            public CellColumn TopMargin { get; set; }
            public CellColumn DefaultTabStop { get; set; }
            public CellColumn TextBkgnd { get; set; }
            public CellColumn TextBkgndTrans { get; set; }
            public CellColumn TextDirection { get; set; }
            public CellColumn VerticalAlign { get; set; }
            public CellColumn TxtWidth { get; set; }
            public CellColumn TxtHeight { get; set; }
            public CellColumn TxtPinX { get; set; }
            public CellColumn TxtPinY { get; set; }
            public CellColumn TxtLocPinX { get; set; }
            public CellColumn TxtLocPinY { get; set; }
            public CellColumn TxtAngle { get; set; }

            public TextBlockFormatCellQuery() :
                base()
            {
                BottomMargin = this.AddCell(VA.ShapeSheet.SRCConstants.BottomMargin);
                LeftMargin = this.AddCell(VA.ShapeSheet.SRCConstants.LeftMargin);
                RightMargin = this.AddCell(VA.ShapeSheet.SRCConstants.RightMargin);
                TopMargin = this.AddCell(VA.ShapeSheet.SRCConstants.TopMargin);
                DefaultTabStop = this.AddCell(VA.ShapeSheet.SRCConstants.DefaultTabStop);
                TextBkgnd = this.AddCell(VA.ShapeSheet.SRCConstants.TextBkgnd);
                TextBkgndTrans = this.AddCell(VA.ShapeSheet.SRCConstants.TextBkgndTrans);
                TextDirection = this.AddCell(VA.ShapeSheet.SRCConstants.TextDirection);
                VerticalAlign = this.AddCell(VA.ShapeSheet.SRCConstants.VerticalAlign);
                TxtPinX = this.AddCell(VA.ShapeSheet.SRCConstants.TxtPinX);
                TxtPinY = this.AddCell(VA.ShapeSheet.SRCConstants.TxtPinY);
                TxtLocPinX = this.AddCell(VA.ShapeSheet.SRCConstants.TxtLocPinX);
                TxtLocPinY = this.AddCell(VA.ShapeSheet.SRCConstants.TxtLocPinY);
                TxtWidth = this.AddCell(VA.ShapeSheet.SRCConstants.TxtWidth);
                TxtHeight = this.AddCell(VA.ShapeSheet.SRCConstants.TxtHeight);
                TxtAngle = this.AddCell(VA.ShapeSheet.SRCConstants.TxtAngle);
            }

            public TextCells GetCells(IList<VA.ShapeSheet.CellData<double>> row)
            {
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