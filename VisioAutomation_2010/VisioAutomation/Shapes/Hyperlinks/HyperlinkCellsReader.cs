using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Hyperlinks
{
    class HyperlinkCellsReader : MultiRowReader<Shapes.Hyperlinks.HyperlinkCells>
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

            this.Address = sec.AddCell(ShapeSheet.SRCConstants.Hyperlink_Address , nameof(ShapeSheet.SRCConstants.Hyperlink_Address));
            this.Default = sec.AddCell(ShapeSheet.SRCConstants.Hyperlink_Default, nameof(ShapeSheet.SRCConstants.Hyperlink_Default));
            this.Description= sec.AddCell(ShapeSheet.SRCConstants.Hyperlink_Description, nameof(ShapeSheet.SRCConstants.Hyperlink_Description));
            this.ExtraInfo= sec.AddCell(ShapeSheet.SRCConstants.Hyperlink_ExtraInfo, nameof(ShapeSheet.SRCConstants.Hyperlink_ExtraInfo));
            this.Frame= sec.AddCell(ShapeSheet.SRCConstants.Hyperlink_Frame, nameof(ShapeSheet.SRCConstants.Hyperlink_Frame));
            this.Invisible= sec.AddCell(ShapeSheet.SRCConstants.Hyperlink_Invisible, nameof(ShapeSheet.SRCConstants.Hyperlink_Invisible));
            this.NewWindow= sec.AddCell(ShapeSheet.SRCConstants.Hyperlink_NewWindow, nameof(ShapeSheet.SRCConstants.Hyperlink_NewWindow));
            this.SortKey= sec.AddCell(ShapeSheet.SRCConstants.Hyperlink_SortKey, nameof(ShapeSheet.SRCConstants.Hyperlink_SortKey));
            this.SubAddress= sec.AddCell(ShapeSheet.SRCConstants.Hyperlink_SubAddress, nameof(ShapeSheet.SRCConstants.Hyperlink_SubAddress));
        }

        public override Shapes.Hyperlinks.HyperlinkCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Shapes.Hyperlinks.HyperlinkCells();

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
 