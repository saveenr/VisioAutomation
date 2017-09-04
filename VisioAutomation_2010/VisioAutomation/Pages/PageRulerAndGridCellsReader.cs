using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    class PageRulerAndGridCellsReader : ReaderSingleRow<VisioAutomation.Pages.PageRulerAndGridCells>
    {
        public CellColumn XGridDensity { get; set; }
        public CellColumn XGridOrigin { get; set; }
        public CellColumn XGridSpacing { get; set; }
        public CellColumn XRulerDensity { get; set; }
        public CellColumn XRulerOrigin { get; set; }
        public CellColumn YGridDensity { get; set; }
        public CellColumn YGridOrigin { get; set; }
        public CellColumn YGridSpacing { get; set; }
        public CellColumn YRulerDensity { get; set; }
        public CellColumn YRulerOrigin { get; set; }

        public PageRulerAndGridCellsReader()
        {
            this.XGridDensity = this.query.AddColumn(SrcConstants.XGridDensity, nameof(SrcConstants.XGridDensity));
            this.XGridOrigin = this.query.AddColumn(SrcConstants.XGridOrigin, nameof(SrcConstants.XGridOrigin));
            this.XGridSpacing = this.query.AddColumn(SrcConstants.XGridSpacing, nameof(SrcConstants.XGridSpacing));
            this.XRulerDensity = this.query.AddColumn(SrcConstants.XRulerDensity, nameof(SrcConstants.XRulerDensity));
            this.XRulerOrigin = this.query.AddColumn(SrcConstants.XRulerOrigin, nameof(SrcConstants.XRulerOrigin));
            this.YGridDensity = this.query.AddColumn(SrcConstants.YGridDensity, nameof(SrcConstants.YGridDensity));
            this.YGridOrigin = this.query.AddColumn(SrcConstants.YGridOrigin, nameof(SrcConstants.YGridOrigin));
            this.YGridSpacing = this.query.AddColumn(SrcConstants.YGridSpacing, nameof(SrcConstants.YGridSpacing));
            this.YRulerDensity = this.query.AddColumn(SrcConstants.YRulerDensity, nameof(SrcConstants.YRulerDensity));
            this.YRulerOrigin = this.query.AddColumn(SrcConstants.YRulerOrigin, nameof(SrcConstants.YRulerOrigin));
        }

        public override PageRulerAndGridCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Pages.PageRulerAndGridCells();
            cells.XGridDensity = row[this.XGridDensity];
            cells.XGridOrigin = row[this.XGridOrigin];
            cells.XGridSpacing = row[this.XGridSpacing];
            cells.XRulerDensity = row[this.XRulerDensity];
            cells.XRulerOrigin = row[this.XRulerOrigin];
            cells.YGridDensity = row[this.YGridDensity];
            cells.YGridOrigin = row[this.YGridOrigin];
            cells.YGridSpacing = row[this.YGridSpacing];
            cells.YRulerDensity = row[this.YRulerDensity];
            cells.YRulerOrigin = row[this.YRulerOrigin];
            return cells;
        }
    }
}