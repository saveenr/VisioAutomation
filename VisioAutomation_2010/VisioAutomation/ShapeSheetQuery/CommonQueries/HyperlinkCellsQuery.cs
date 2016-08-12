using VisioAutomation.ShapeSheetQuery.Columns;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery.CommonQueries
{
    class HyperlinkCellsQuery : Query
    {

        public ColumnCellIndex Address { get; set; }
        public ColumnCellIndex Description { get; set; }
        public ColumnCellIndex ExtraInfo { get; set; }
        public ColumnCellIndex Frame { get; set; }
        public ColumnCellIndex SortKey { get; set; }
        public ColumnCellIndex SubAddress { get; set; }
        public ColumnCellIndex NewWindow { get; set; }
        public ColumnCellIndex Default { get; set; }
        public ColumnCellIndex Invisible { get; set; }

        public HyperlinkCellsQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionHyperlink);

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

        public Shapes.Hyperlinks.HyperlinkCells GetCells(ShapeSheet.CellData<double>[] row)
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
 