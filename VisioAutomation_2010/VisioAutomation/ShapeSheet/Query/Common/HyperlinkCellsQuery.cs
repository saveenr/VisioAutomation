using System.Windows.Forms;

namespace VisioAutomation.ShapeSheet.Query.Common
{
    class HyperlinkCellsQuery : CellQuery
    {

        public Query.CellColumn Address { get; set; }
        public Query.CellColumn Description { get; set; }
        public Query.CellColumn ExtraInfo { get; set; }
        public Query.CellColumn Frame { get; set; }
        public Query.CellColumn SortKey { get; set; }
        public Query.CellColumn SubAddress { get; set; }

        public Query.CellColumn NewWindow { get; set; }
        public Query.CellColumn Default { get; set; }
        public Query.CellColumn Invisible { get; set; }

        public HyperlinkCellsQuery()
        {
            var sec = this.AddSection(Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionHyperlink);

            this.Address = sec.AddCell(SRCConstants.Hyperlink_Address , nameof(SRCConstants.Hyperlink_Address));
            this.Default = sec.AddCell(SRCConstants.Hyperlink_Default, nameof(SRCConstants.Hyperlink_Default));
            this.Description= sec.AddCell(SRCConstants.Hyperlink_Description, nameof(SRCConstants.Hyperlink_Description));
            this.ExtraInfo= sec.AddCell(SRCConstants.Hyperlink_ExtraInfo, nameof(SRCConstants.Hyperlink_ExtraInfo));
            this.Frame= sec.AddCell(SRCConstants.Hyperlink_Frame, nameof(SRCConstants.Hyperlink_Frame));
            this.Invisible= sec.AddCell(SRCConstants.Hyperlink_Invisible, nameof(SRCConstants.Hyperlink_Invisible));
            this.NewWindow= sec.AddCell(SRCConstants.Hyperlink_NewWindow, nameof(SRCConstants.Hyperlink_NewWindow));
            this.SortKey= sec.AddCell(SRCConstants.Hyperlink_SortKey, nameof(SRCConstants.Hyperlink_SortKey));
            this.SubAddress= sec.AddCell(SRCConstants.Hyperlink_SubAddress, nameof(SRCConstants.Hyperlink_SubAddress));
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
 