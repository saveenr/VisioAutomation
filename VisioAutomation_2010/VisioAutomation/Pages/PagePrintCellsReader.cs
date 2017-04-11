using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    class PagePrintCellsReader : SingleRowReader<VisioAutomation.Pages.PagePrintCells>
    {
        public CellColumn PrintLeftMargin { get; set; }
        public CellColumn PrintCenterX { get; set; }
        public CellColumn PrintCenterY { get; set; }
        public CellColumn PrintOnPage { get; set; }
        public CellColumn PrintBottomMargin { get; set; }
        public CellColumn PrintRightMargin { get; set; }
        public CellColumn PrintPagesX { get; set; }
        public CellColumn PrintPagesY { get; set; }
        public CellColumn PrintTopMargin { get; set; }
        public CellColumn PrintPaperKind { get; set; }
        public CellColumn PrintGrid { get; set; }
        public CellColumn PrintPageOrientation { get; set; }
        public CellColumn PrintScaleX { get; set; }
        public CellColumn PrintScaleY { get; set; }
        public CellColumn PrintPaperSource { get; set; }

        public PagePrintCellsReader()
        {
            this.PrintLeftMargin = this.query.AddCell(SrcConstants.PrintLeftMargin, nameof(SrcConstants.PrintLeftMargin));
            this.PrintCenterX = this.query.AddCell(SrcConstants.PrintCenterX, nameof(SrcConstants.PrintCenterX));
            this.PrintCenterY = this.query.AddCell(SrcConstants.PrintCenterY, nameof(SrcConstants.PrintCenterY));
            this.PrintOnPage = this.query.AddCell(SrcConstants.PrintOnPage, nameof(SrcConstants.PrintOnPage));
            this.PrintBottomMargin = this.query.AddCell(SrcConstants.PrintBottomMargin, nameof(SrcConstants.PrintBottomMargin));
            this.PrintRightMargin = this.query.AddCell(SrcConstants.PrintRightMargin, nameof(SrcConstants.PrintRightMargin));
            this.PrintPagesX = this.query.AddCell(SrcConstants.PrintPagesX, nameof(SrcConstants.PrintPagesX));
            this.PrintPagesY = this.query.AddCell(SrcConstants.PrintPagesY, nameof(SrcConstants.PrintPagesY));
            this.PrintTopMargin = this.query.AddCell(SrcConstants.PrintTopMargin, nameof(SrcConstants.PrintTopMargin));
            this.PrintPaperKind = this.query.AddCell(SrcConstants.PrintPaperKind, nameof(SrcConstants.PrintPaperKind));
            this.PrintGrid = this.query.AddCell(SrcConstants.PrintGrid, nameof(SrcConstants.PrintGrid));
            this.PrintPageOrientation = this.query.AddCell(SrcConstants.PrintPageOrientation, nameof(SrcConstants.PrintPageOrientation));
            this.PrintScaleX = this.query.AddCell(SrcConstants.PrintScaleX, nameof(SrcConstants.PrintScaleX));
            this.PrintScaleY = this.query.AddCell(SrcConstants.PrintScaleY, nameof(SrcConstants.PrintScaleY));
            this.PrintPaperSource = this.query.AddCell(SrcConstants.PrintPaperSource, nameof(SrcConstants.PrintPaperSource));
        }

        public override Pages.PagePrintCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Pages.PagePrintCells();
            cells.LeftMargin = row[this.PrintLeftMargin];
            cells.CenterX = row[this.PrintCenterX];
            cells.CenterY = row[this.PrintCenterY];
            cells.OnPage = row[this.PrintOnPage];
            cells.BottomMargin = row[this.PrintBottomMargin];
            cells.RightMargin = row[this.PrintRightMargin];
            cells.PagesX = row[this.PrintPagesX];
            cells.PagesY = row[this.PrintPagesY];
            cells.TopMargin = row[this.PrintTopMargin];
            cells.PaperKind = row[this.PrintPaperKind];
            cells.Grid = row[this.PrintGrid];
            cells.Orientation = row[this.PrintPageOrientation];
            cells.ScaleX = row[this.PrintScaleX];
            cells.ScaleY = row[this.PrintScaleY];
            cells.PaperSource = row[this.PrintPaperSource];
            return cells;
        }
    }
}