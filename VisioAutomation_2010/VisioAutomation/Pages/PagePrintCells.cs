using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    public class PagePrintCells : CellGroupSingleRow
    {
        public CellValueLiteral LeftMargin { get; set; }
        public CellValueLiteral CenterX { get; set; }
        public CellValueLiteral CenterY { get; set; }
        public CellValueLiteral OnPage { get; set; }
        public CellValueLiteral BottomMargin { get; set; }
        public CellValueLiteral RightMargin { get; set; }
        public CellValueLiteral PagesX { get; set; }
        public CellValueLiteral PagesY { get; set; }
        public CellValueLiteral TopMargin { get; set; }
        public CellValueLiteral PaperKind { get; set; }
        public CellValueLiteral Grid { get; set; }
        public CellValueLiteral Orientation { get; set; }
        public CellValueLiteral ScaleX { get; set; }
        public CellValueLiteral ScaleY { get; set; }
        public CellValueLiteral PaperSource { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.PrintLeftMargin, this.LeftMargin);
                yield return SrcValuePair.Create(SrcConstants.PrintCenterX, this.CenterX);
                yield return SrcValuePair.Create(SrcConstants.PrintCenterY, this.CenterY);
                yield return SrcValuePair.Create(SrcConstants.PrintOnPage, this.OnPage);
                yield return SrcValuePair.Create(SrcConstants.PrintBottomMargin, this.BottomMargin);
                yield return SrcValuePair.Create(SrcConstants.PrintRightMargin, this.RightMargin);
                yield return SrcValuePair.Create(SrcConstants.PrintPagesX, this.PagesX);
                yield return SrcValuePair.Create(SrcConstants.PrintPagesY, this.PagesY);
                yield return SrcValuePair.Create(SrcConstants.PrintTopMargin, this.TopMargin);
                yield return SrcValuePair.Create(SrcConstants.PrintPaperKind, this.PaperKind);
                yield return SrcValuePair.Create(SrcConstants.PrintGrid, this.Grid);
                yield return SrcValuePair.Create(SrcConstants.PrintPageOrientation, this.Orientation);
                yield return SrcValuePair.Create(SrcConstants.PrintScaleX, this.ScaleX);
                yield return SrcValuePair.Create(SrcConstants.PrintScaleY, this.ScaleY);
                yield return SrcValuePair.Create(SrcConstants.PrintPaperSource, this.PaperSource);
            }
        }

        public static PagePrintCells GetCells(Microsoft.Office.Interop.Visio.Shape shape, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<PagePrintCellsReader> lazy_query = new System.Lazy<PagePrintCellsReader>();

        class PagePrintCellsReader : ReaderSingleRow<PagePrintCells>
        {
            public CellColumn LeftMargin { get; set; }
            public CellColumn CenterX { get; set; }
            public CellColumn CenterY { get; set; }
            public CellColumn OnPage { get; set; }
            public CellColumn BottomMargin { get; set; }
            public CellColumn RightMargin { get; set; }
            public CellColumn PagesX { get; set; }
            public CellColumn PagesY { get; set; }
            public CellColumn TopMargin { get; set; }
            public CellColumn PaperKind { get; set; }
            public CellColumn Grid { get; set; }
            public CellColumn PageOrientation { get; set; }
            public CellColumn ScaleX { get; set; }
            public CellColumn ScaleY { get; set; }
            public CellColumn PaperSource { get; set; }

            public PagePrintCellsReader()
            {
                this.LeftMargin = this.query.Columns.Add(SrcConstants.PrintLeftMargin, nameof(this.LeftMargin));
                this.CenterX = this.query.Columns.Add(SrcConstants.PrintCenterX, nameof(this.CenterX));
                this.CenterY = this.query.Columns.Add(SrcConstants.PrintCenterY, nameof(this.CenterY));
                this.OnPage = this.query.Columns.Add(SrcConstants.PrintOnPage, nameof(this.OnPage));
                this.BottomMargin = this.query.Columns.Add(SrcConstants.PrintBottomMargin, nameof(this.BottomMargin));
                this.RightMargin = this.query.Columns.Add(SrcConstants.PrintRightMargin, nameof(this.RightMargin));
                this.PagesX = this.query.Columns.Add(SrcConstants.PrintPagesX, nameof(this.PagesX));
                this.PagesY = this.query.Columns.Add(SrcConstants.PrintPagesY, nameof(this.PagesY));
                this.TopMargin = this.query.Columns.Add(SrcConstants.PrintTopMargin, nameof(this.TopMargin));
                this.PaperKind = this.query.Columns.Add(SrcConstants.PrintPaperKind, nameof(this.PaperKind));
                this.Grid = this.query.Columns.Add(SrcConstants.PrintGrid, nameof(this.Grid));
                this.PageOrientation = this.query.Columns.Add(SrcConstants.PrintPageOrientation, nameof(this.PageOrientation));
                this.ScaleX = this.query.Columns.Add(SrcConstants.PrintScaleX, nameof(this.ScaleX));
                this.ScaleY = this.query.Columns.Add(SrcConstants.PrintScaleY, nameof(this.ScaleY));
                this.PaperSource = this.query.Columns.Add(SrcConstants.PrintPaperSource, nameof(this.PaperSource));
            }

            public override PagePrintCells CellDataToCellGroup(Utilities.ArraySegment<string> row)
            {
                var cells = new PagePrintCells();
                cells.LeftMargin = row[this.LeftMargin];
                cells.CenterX = row[this.CenterX];
                cells.CenterY = row[this.CenterY];
                cells.OnPage = row[this.OnPage];
                cells.BottomMargin = row[this.BottomMargin];
                cells.RightMargin = row[this.RightMargin];
                cells.PagesX = row[this.PagesX];
                cells.PagesY = row[this.PagesY];
                cells.TopMargin = row[this.TopMargin];
                cells.PaperKind = row[this.PaperKind];
                cells.Grid = row[this.Grid];
                cells.Orientation = row[this.PageOrientation];
                cells.ScaleX = row[this.ScaleX];
                cells.ScaleY = row[this.ScaleY];
                cells.PaperSource = row[this.PaperSource];
                return cells;
            }
        }

    }
}