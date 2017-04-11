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
        public CellColumn PageDrawingScale { get; set; }
        public CellColumn PageDrawingScaleType { get; set; }
        public CellColumn PageDrawingSizeType { get; set; }
        public CellColumn PageInhibitSnap { get; set; }
        public CellColumn PageHeight { get; set; }
        public CellColumn PageScale { get; set; }
        public CellColumn PageWidth { get; set; }

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

            this.PageDrawingScale = this.query.AddCell(SrcConstants.PageDrawingScale, nameof(SrcConstants.PageDrawingScale));
            this.PageDrawingScaleType = this.query.AddCell(SrcConstants.PageDrawingScaleType, nameof(SrcConstants.PageDrawingScaleType));
            this.PageDrawingSizeType = this.query.AddCell(SrcConstants.PageDrawingSizeType, nameof(SrcConstants.PageDrawingSizeType));
            this.PageInhibitSnap = this.query.AddCell(SrcConstants.PageInhibitSnap, nameof(SrcConstants.PageInhibitSnap));
            this.PageHeight = this.query.AddCell(SrcConstants.PageHeight, nameof(SrcConstants.PageHeight));
            this.PageScale = this.query.AddCell(SrcConstants.PageScale, nameof(SrcConstants.PageScale));
            this.PageWidth = this.query.AddCell(SrcConstants.PageWidth, nameof(SrcConstants.PageWidth));
        }

        public override Pages.PagePrintCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Pages.PagePrintCells();
            cells.PrintLeftMargin = row[this.PrintLeftMargin];
            cells.PrintCenterX = row[this.PrintCenterX];
            cells.PrintCenterY = row[this.PrintCenterY];
            cells.PrintOnPage = row[this.PrintOnPage];
            cells.PrintBottomMargin = row[this.PrintBottomMargin];
            cells.PrintRightMargin = row[this.PrintRightMargin];
            cells.PrintPagesX = row[this.PrintPagesX];
            cells.PrintPagesY = row[this.PrintPagesY];
            cells.PrintTopMargin = row[this.PrintTopMargin];
            cells.PrintPaperKind = row[this.PrintPaperKind];
            cells.PrintGrid = row[this.PrintGrid];
            cells.PrintPageOrientation = row[this.PrintPageOrientation];
            cells.PrintScaleX = row[this.PrintScaleX];
            cells.PrintScaleY = row[this.PrintScaleY];
            cells.PrintPaperSource = row[this.PrintPaperSource];
            cells.PageDrawingScale = row[this.PageDrawingScale];
            cells.PageDrawingScaleType = row[this.PageDrawingScaleType];
            cells.PageDrawingSizeType = row[this.PageDrawingSizeType];
            cells.PageInhibitSnap = row[this.PageInhibitSnap];
            cells.PageHeight = row[this.PageHeight];
            cells.PageScale = row[this.PageScale];
            cells.PageWidth = row[this.PageWidth];
            return cells;
        }
    }
}