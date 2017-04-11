using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    class PageCellsReader : SingleRowReader<VisioAutomation.Pages.PageCells>
    {
        public CellColumn PageDrawingScale { get; set; }
        public CellColumn PageDrawingScaleType { get; set; }
        public CellColumn PageDrawingSizeType { get; set; }
        public CellColumn PageInhibitSnap { get; set; }
        public CellColumn PageHeight { get; set; }
        public CellColumn PageScale { get; set; }
        public CellColumn PageWidth { get; set; }
        public CellColumn PageShadowObliqueAngle { get; set; }
        public CellColumn PageShadowOffsetX { get; set; }
        public CellColumn PageShadowOffsetY { get; set; }
        public CellColumn PageShadowScaleFactor { get; set; }
        public CellColumn PageShadowType { get; set; }
        public CellColumn PageUIVisibility { get; set; }
        public CellColumn PageDrawingResizeType { get; set; }

        public PageCellsReader()
        {
            this.PageDrawingScale = this.query.AddCell(SrcConstants.PageDrawingScale, nameof(SrcConstants.PageDrawingScale));
            this.PageDrawingScaleType = this.query.AddCell(SrcConstants.PageDrawingScaleType, nameof(SrcConstants.PageDrawingScaleType));
            this.PageDrawingSizeType = this.query.AddCell(SrcConstants.PageDrawingSizeType, nameof(SrcConstants.PageDrawingSizeType));
            this.PageInhibitSnap = this.query.AddCell(SrcConstants.PageInhibitSnap, nameof(SrcConstants.PageInhibitSnap));
            this.PageHeight = this.query.AddCell(SrcConstants.PageHeight, nameof(SrcConstants.PageHeight));
            this.PageScale = this.query.AddCell(SrcConstants.PageScale, nameof(SrcConstants.PageScale));
            this.PageWidth = this.query.AddCell(SrcConstants.PageWidth, nameof(SrcConstants.PageWidth));
            this.PageShadowObliqueAngle = this.query.AddCell(SrcConstants.PageShadowObliqueAngle, nameof(SrcConstants.PageShadowObliqueAngle));
            this.PageShadowOffsetX = this.query.AddCell(SrcConstants.PageShadowOffsetX, nameof(SrcConstants.PageShadowOffsetX));
            this.PageShadowOffsetY = this.query.AddCell(SrcConstants.PageShadowOffsetY, nameof(SrcConstants.PageShadowOffsetY));
            this.PageShadowScaleFactor = this.query.AddCell(SrcConstants.PageShadowScaleFactor, nameof(SrcConstants.PageShadowScaleFactor));
            this.PageShadowType = this.query.AddCell(SrcConstants.PageShadowType, nameof(SrcConstants.PageShadowType));
            this.PageUIVisibility = this.query.AddCell(SrcConstants.PageUIVisibility, nameof(SrcConstants.PageUIVisibility));
            this.PageDrawingResizeType = this.query.AddCell(SrcConstants.PageDrawingResizeType, nameof(SrcConstants.PageDrawingResizeType));
        }

        public override Pages.PageCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Pages.PageCells();
            cells.PageDrawingScale = row[this.PageDrawingScale];
            cells.PageDrawingScaleType = row[this.PageDrawingScaleType];
            cells.PageDrawingSizeType = row[this.PageDrawingSizeType];
            cells.PageInhibitSnap = row[this.PageInhibitSnap];
            cells.PageHeight = row[this.PageHeight];
            cells.PageScale = row[this.PageScale];
            cells.PageWidth = row[this.PageWidth];
            cells.PageShadowObliqueAngle = row[this.PageShadowObliqueAngle];
            cells.PageShadowOffsetX = row[this.PageShadowOffsetX];
            cells.PageShadowOffsetY = row[this.PageShadowOffsetY];
            cells.PageShadowScaleFactor = row[this.PageShadowScaleFactor];
            cells.PageShadowType = row[this.PageShadowType];
            cells.PageUIVisibility = row[this.PageUIVisibility];
            cells.PageDrawingResizeType = row[this.PageDrawingResizeType];
            return cells;
        }
    }
}