using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery.Common
{
    class HyperlinkCellsQuery : CellQuery
    {

        public VisioAutomation.ShapeSheetQuery.CellColumn Address { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Description { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn ExtraInfo { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Frame { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn SortKey { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn SubAddress { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn NewWindow { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Default { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Invisible { get; set; }

        public HyperlinkCellsQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionHyperlink);

            this.Address = sec.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_Address , nameof(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_Address));
            this.Default = sec.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_Default, nameof(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_Default));
            this.Description= sec.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_Description, nameof(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_Description));
            this.ExtraInfo= sec.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_ExtraInfo, nameof(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_ExtraInfo));
            this.Frame= sec.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_Frame, nameof(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_Frame));
            this.Invisible= sec.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_Invisible, nameof(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_Invisible));
            this.NewWindow= sec.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_NewWindow, nameof(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_NewWindow));
            this.SortKey= sec.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_SortKey, nameof(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_SortKey));
            this.SubAddress= sec.AddCell(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_SubAddress, nameof(VisioAutomation.ShapeSheet.SRCConstants.Hyperlink_SubAddress));
        }

        public VisioAutomation.Shapes.Hyperlinks.HyperlinkCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Shapes.Hyperlinks.HyperlinkCells();

            // cells.X = Extensions.CellDataMethods.ToInt(row[this.X]);

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
 