using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    class HyperlinkCellsReader : ReaderMultiRow<HyperlinkCells>
    {

        public SubQueryColumn Address { get; set; }
        public SubQueryColumn Description { get; set; }
        public SubQueryColumn ExtraInfo { get; set; }
        public SubQueryColumn Frame { get; set; }
        public SubQueryColumn SortKey { get; set; }
        public SubQueryColumn SubAddress { get; set; }
        public SubQueryColumn NewWindow { get; set; }
        public SubQueryColumn Default { get; set; }
        public SubQueryColumn Invisible { get; set; }

        public HyperlinkCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionHyperlink);

            this.Address = sec.AddCell(ShapeSheet.SrcConstants.HyperlinkAddress , nameof(ShapeSheet.SrcConstants.HyperlinkAddress));
            this.Default = sec.AddCell(ShapeSheet.SrcConstants.HyperlinkDefault, nameof(ShapeSheet.SrcConstants.HyperlinkDefault));
            this.Description= sec.AddCell(ShapeSheet.SrcConstants.HyperlinkDescription, nameof(ShapeSheet.SrcConstants.HyperlinkDescription));
            this.ExtraInfo= sec.AddCell(ShapeSheet.SrcConstants.HyperlinkExtraInfo, nameof(ShapeSheet.SrcConstants.HyperlinkExtraInfo));
            this.Frame= sec.AddCell(ShapeSheet.SrcConstants.HyperlinkFrame, nameof(ShapeSheet.SrcConstants.HyperlinkFrame));
            this.Invisible= sec.AddCell(ShapeSheet.SrcConstants.HyperlinkInvisible, nameof(ShapeSheet.SrcConstants.HyperlinkInvisible));
            this.NewWindow= sec.AddCell(ShapeSheet.SrcConstants.HyperlinkNewWindow, nameof(ShapeSheet.SrcConstants.HyperlinkNewWindow));
            this.SortKey= sec.AddCell(ShapeSheet.SrcConstants.HyperlinkSortKey, nameof(ShapeSheet.SrcConstants.HyperlinkSortKey));
            this.SubAddress= sec.AddCell(ShapeSheet.SrcConstants.HyperlinkSubAddress, nameof(ShapeSheet.SrcConstants.HyperlinkSubAddress));
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
 