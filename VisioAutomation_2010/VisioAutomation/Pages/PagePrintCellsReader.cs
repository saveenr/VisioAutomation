using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    class PagePrintCellsReader : SingleRowReader<VisioAutomation.Pages.PagePrintCells>
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
            this.LeftMargin = this.query.AddCell(SrcConstants.PrintLeftMargin, nameof(SrcConstants.PrintLeftMargin));
            this.CenterX = this.query.AddCell(SrcConstants.PrintCenterX, nameof(SrcConstants.PrintCenterX));
            this.CenterY = this.query.AddCell(SrcConstants.PrintCenterY, nameof(SrcConstants.PrintCenterY));
            this.OnPage = this.query.AddCell(SrcConstants.PrintOnPage, nameof(SrcConstants.PrintOnPage));
            this.BottomMargin = this.query.AddCell(SrcConstants.PrintBottomMargin, nameof(SrcConstants.PrintBottomMargin));
            this.RightMargin = this.query.AddCell(SrcConstants.PrintRightMargin, nameof(SrcConstants.PrintRightMargin));
            this.PagesX = this.query.AddCell(SrcConstants.PrintPagesX, nameof(SrcConstants.PrintPagesX));
            this.PagesY = this.query.AddCell(SrcConstants.PrintPagesY, nameof(SrcConstants.PrintPagesY));
            this.TopMargin = this.query.AddCell(SrcConstants.PrintTopMargin, nameof(SrcConstants.PrintTopMargin));
            this.PaperKind = this.query.AddCell(SrcConstants.PrintPaperKind, nameof(SrcConstants.PrintPaperKind));
            this.Grid = this.query.AddCell(SrcConstants.PrintGrid, nameof(SrcConstants.PrintGrid));
            this.PageOrientation = this.query.AddCell(SrcConstants.PrintPageOrientation, nameof(SrcConstants.PrintPageOrientation));
            this.ScaleX = this.query.AddCell(SrcConstants.PrintScaleX, nameof(SrcConstants.PrintScaleX));
            this.ScaleY = this.query.AddCell(SrcConstants.PrintScaleY, nameof(SrcConstants.PrintScaleY));
            this.PaperSource = this.query.AddCell(SrcConstants.PrintPaperSource, nameof(SrcConstants.PrintPaperSource));
        }

        public override Pages.PagePrintCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
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