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

            this.Address = sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkAddress , nameof(ShapeSheet.SrcConstants.HyperlinkAddress));
            this.Default = sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkDefault, nameof(ShapeSheet.SrcConstants.HyperlinkDefault));
            this.Description= sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkDescription, nameof(ShapeSheet.SrcConstants.HyperlinkDescription));
            this.ExtraInfo= sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkExtraInfo, nameof(ShapeSheet.SrcConstants.HyperlinkExtraInfo));
            this.Frame= sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkFrame, nameof(ShapeSheet.SrcConstants.HyperlinkFrame));
            this.Invisible= sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkInvisible, nameof(ShapeSheet.SrcConstants.HyperlinkInvisible));
            this.NewWindow= sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkNewWindow, nameof(ShapeSheet.SrcConstants.HyperlinkNewWindow));
            this.SortKey= sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkSortKey, nameof(ShapeSheet.SrcConstants.HyperlinkSortKey));
            this.SubAddress= sec.Columns.Add(ShapeSheet.SrcConstants.HyperlinkSubAddress, nameof(ShapeSheet.SrcConstants.HyperlinkSubAddress));
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
 