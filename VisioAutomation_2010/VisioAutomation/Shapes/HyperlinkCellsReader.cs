using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    class HyperlinkCellsReader : ReaderMultiRow<HyperlinkCells>
    {

        public SectionQueryColumn Address { get; set; }
        public SectionQueryColumn Description { get; set; }
        public SectionQueryColumn ExtraInfo { get; set; }
        public SectionQueryColumn Frame { get; set; }
        public SectionQueryColumn SortKey { get; set; }
        public SectionQueryColumn SubAddress { get; set; }
        public SectionQueryColumn NewWindow { get; set; }
        public SectionQueryColumn Default { get; set; }
        public SectionQueryColumn Invisible { get; set; }

        public HyperlinkCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionHyperlink);

            this.Address = sec.AddColumn(ShapeSheet.SrcConstants.HyperlinkAddress , nameof(ShapeSheet.SrcConstants.HyperlinkAddress));
            this.Default = sec.AddColumn(ShapeSheet.SrcConstants.HyperlinkDefault, nameof(ShapeSheet.SrcConstants.HyperlinkDefault));
            this.Description= sec.AddColumn(ShapeSheet.SrcConstants.HyperlinkDescription, nameof(ShapeSheet.SrcConstants.HyperlinkDescription));
            this.ExtraInfo= sec.AddColumn(ShapeSheet.SrcConstants.HyperlinkExtraInfo, nameof(ShapeSheet.SrcConstants.HyperlinkExtraInfo));
            this.Frame= sec.AddColumn(ShapeSheet.SrcConstants.HyperlinkFrame, nameof(ShapeSheet.SrcConstants.HyperlinkFrame));
            this.Invisible= sec.AddColumn(ShapeSheet.SrcConstants.HyperlinkInvisible, nameof(ShapeSheet.SrcConstants.HyperlinkInvisible));
            this.NewWindow= sec.AddColumn(ShapeSheet.SrcConstants.HyperlinkNewWindow, nameof(ShapeSheet.SrcConstants.HyperlinkNewWindow));
            this.SortKey= sec.AddColumn(ShapeSheet.SrcConstants.HyperlinkSortKey, nameof(ShapeSheet.SrcConstants.HyperlinkSortKey));
            this.SubAddress= sec.AddColumn(ShapeSheet.SrcConstants.HyperlinkSubAddress, nameof(ShapeSheet.SrcConstants.HyperlinkSubAddress));
        }

        public override HyperlinkCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new HyperlinkCells();

            cells.Address = row[this.Address].Formula;
            cells.Description= row[this.Description].Formula;
            cells.ExtraInfo= row[this.ExtraInfo].Formula;
            cells.Frame= row[this.Frame].Formula;
            cells.SortKey= row[this.SortKey].Formula;
            cells.SubAddress= row[this.SubAddress].Formula;
            cells.NewWindow = row[this.NewWindow];
            cells.Default = row[this.Default];
            cells.Invisible = row[this.Invisible];

            return cells;
        }
    }
}
 