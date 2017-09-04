using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    class PageFormatCellsReader : ReaderSingleRow<VisioAutomation.Pages.PageFormatCells>
    {
        public CellColumn DrawingScale { get; set; }
        public CellColumn DrawingScaleType { get; set; }
        public CellColumn DrawingSizeType { get; set; }
        public CellColumn InhibitSnap { get; set; }
        public CellColumn Height { get; set; }
        public CellColumn Scale { get; set; }
        public CellColumn Width { get; set; }
        public CellColumn ShadowObliqueAngle { get; set; }
        public CellColumn ShadowOffsetX { get; set; }
        public CellColumn ShadowOffsetY { get; set; }
        public CellColumn ShadowScaleFactor { get; set; }
        public CellColumn ShadowType { get; set; }
        public CellColumn UIVisibility { get; set; }
        public CellColumn DrawingResizeType { get; set; }

        public PageFormatCellsReader()
        {
            this.DrawingScale = this.query.AddColumn(SrcConstants.PageDrawingScale, nameof(SrcConstants.PageDrawingScale));
            this.DrawingScaleType = this.query.AddColumn(SrcConstants.PageDrawingScaleType, nameof(SrcConstants.PageDrawingScaleType));
            this.DrawingSizeType = this.query.AddColumn(SrcConstants.PageDrawingSizeType, nameof(SrcConstants.PageDrawingSizeType));
            this.InhibitSnap = this.query.AddColumn(SrcConstants.PageInhibitSnap, nameof(SrcConstants.PageInhibitSnap));
            this.Height = this.query.AddColumn(SrcConstants.PageHeight, nameof(SrcConstants.PageHeight));
            this.Scale = this.query.AddColumn(SrcConstants.PageScale, nameof(SrcConstants.PageScale));
            this.Width = this.query.AddColumn(SrcConstants.PageWidth, nameof(SrcConstants.PageWidth));
            this.ShadowObliqueAngle = this.query.AddColumn(SrcConstants.PageShadowObliqueAngle, nameof(SrcConstants.PageShadowObliqueAngle));
            this.ShadowOffsetX = this.query.AddColumn(SrcConstants.PageShadowOffsetX, nameof(SrcConstants.PageShadowOffsetX));
            this.ShadowOffsetY = this.query.AddColumn(SrcConstants.PageShadowOffsetY, nameof(SrcConstants.PageShadowOffsetY));
            this.ShadowScaleFactor = this.query.AddColumn(SrcConstants.PageShadowScaleFactor, nameof(SrcConstants.PageShadowScaleFactor));
            this.ShadowType = this.query.AddColumn(SrcConstants.PageShadowType, nameof(SrcConstants.PageShadowType));
            this.UIVisibility = this.query.AddColumn(SrcConstants.PageUIVisibility, nameof(SrcConstants.PageUIVisibility));
            this.DrawingResizeType = this.query.AddColumn(SrcConstants.PageDrawingResizeType, nameof(SrcConstants.PageDrawingResizeType));
        }

        public override Pages.PageFormatCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Pages.PageFormatCells();
            cells.DrawingScale = row[this.DrawingScale];
            cells.DrawingScaleType = row[this.DrawingScaleType];
            cells.DrawingSizeType = row[this.DrawingSizeType];
            cells.InhibitSnap = row[this.InhibitSnap];
            cells.Height = row[this.Height];
            cells.Scale = row[this.Scale];
            cells.Width = row[this.Width];
            cells.ShadowObliqueAngle = row[this.ShadowObliqueAngle];
            cells.ShadowOffsetX = row[this.ShadowOffsetX];
            cells.ShadowOffsetY = row[this.ShadowOffsetY];
            cells.ShadowScaleFactor = row[this.ShadowScaleFactor];
            cells.ShadowType = row[this.ShadowType];
            cells.UIVisibility = row[this.UIVisibility];
            cells.DrawingResizeType = row[this.DrawingResizeType];
            return cells;
        }
    }
}