using VisioAutomation.ShapeSheet.Queries.Columns;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Queries.CommonQueries
{
    class HyperlinkCellsQuery : CellGroupMultiRowQuery<Shapes.Hyperlinks.HyperlinkCells, double>
    {

        public ColumnSubQuery Address { get; set; }
        public ColumnSubQuery Description { get; set; }
        public ColumnSubQuery ExtraInfo { get; set; }
        public ColumnSubQuery Frame { get; set; }
        public ColumnSubQuery SortKey { get; set; }
        public ColumnSubQuery SubAddress { get; set; }
        public ColumnSubQuery NewWindow { get; set; }
        public ColumnSubQuery Default { get; set; }
        public ColumnSubQuery Invisible { get; set; }

        public HyperlinkCellsQuery()
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

        public override Shapes.Hyperlinks.HyperlinkCells CellDataToCellGroup(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Shapes.Hyperlinks.HyperlinkCells();

            cells.Address = row[this.Address].Formula;
            cells.Description= row[this.Description].Formula;
            cells.ExtraInfo= row[this.ExtraInfo].Formula;
            cells.Frame= row[this.Frame].Formula;
            cells.SortKey= row[this.SortKey].Formula;
            cells.SubAddress= row[this.SubAddress].Formula;

            cells.NewWindow = Extensions.CellDataMethods.ToBool(row[this.NewWindow]);
            cells.Default = Extensions.CellDataMethods.ToBool(row[this.Default].Formula);
            cells.Invisible = Extensions.CellDataMethods.ToBool(row[this.Invisible].Formula);

            return cells;
        }
    }
}
 