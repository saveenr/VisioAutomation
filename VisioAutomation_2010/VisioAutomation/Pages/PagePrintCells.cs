using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    public class PagePrintCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral LeftMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral CenterX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral CenterY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral OnPage { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral BottomMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral RightMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PagesX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PagesY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral TopMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PaperKind { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Grid { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Orientation { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ScaleX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ScaleY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PaperSource { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintLeftMargin, this.LeftMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintCenterX, this.CenterX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintCenterY, this.CenterY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintOnPage, this.OnPage.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintBottomMargin, this.BottomMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintRightMargin, this.RightMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintPagesX, this.PagesX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintPagesY, this.PagesY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintTopMargin, this.TopMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintPaperKind, this.PaperKind.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintGrid, this.Grid.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintPageOrientation, this.Orientation.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintScaleX, this.ScaleX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintScaleY, this.ScaleY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PrintPaperSource, this.PaperSource.Value);
            }
        }

        public static PagePrintCells GetValues(Microsoft.Office.Interop.Visio.Shape shape, CellValueType cvt)
        {
            var query = PagePrintCells.lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<PagePrintCellsReader> lazy_query = new System.Lazy<PagePrintCellsReader>();

        class PagePrintCellsReader : ReaderSingleRow<VisioAutomation.Pages.PagePrintCells>
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
                this.LeftMargin = this.query.Columns.Add(SrcConstants.PrintLeftMargin, nameof(SrcConstants.PrintLeftMargin));
                this.CenterX = this.query.Columns.Add(SrcConstants.PrintCenterX, nameof(SrcConstants.PrintCenterX));
                this.CenterY = this.query.Columns.Add(SrcConstants.PrintCenterY, nameof(SrcConstants.PrintCenterY));
                this.OnPage = this.query.Columns.Add(SrcConstants.PrintOnPage, nameof(SrcConstants.PrintOnPage));
                this.BottomMargin = this.query.Columns.Add(SrcConstants.PrintBottomMargin, nameof(SrcConstants.PrintBottomMargin));
                this.RightMargin = this.query.Columns.Add(SrcConstants.PrintRightMargin, nameof(SrcConstants.PrintRightMargin));
                this.PagesX = this.query.Columns.Add(SrcConstants.PrintPagesX, nameof(SrcConstants.PrintPagesX));
                this.PagesY = this.query.Columns.Add(SrcConstants.PrintPagesY, nameof(SrcConstants.PrintPagesY));
                this.TopMargin = this.query.Columns.Add(SrcConstants.PrintTopMargin, nameof(SrcConstants.PrintTopMargin));
                this.PaperKind = this.query.Columns.Add(SrcConstants.PrintPaperKind, nameof(SrcConstants.PrintPaperKind));
                this.Grid = this.query.Columns.Add(SrcConstants.PrintGrid, nameof(SrcConstants.PrintGrid));
                this.PageOrientation = this.query.Columns.Add(SrcConstants.PrintPageOrientation, nameof(SrcConstants.PrintPageOrientation));
                this.ScaleX = this.query.Columns.Add(SrcConstants.PrintScaleX, nameof(SrcConstants.PrintScaleX));
                this.ScaleY = this.query.Columns.Add(SrcConstants.PrintScaleY, nameof(SrcConstants.PrintScaleY));
                this.PaperSource = this.query.Columns.Add(SrcConstants.PrintPaperSource, nameof(SrcConstants.PrintPaperSource));
            }

            public override Pages.PagePrintCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
            {
                var cells = new Pages.PagePrintCells();
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