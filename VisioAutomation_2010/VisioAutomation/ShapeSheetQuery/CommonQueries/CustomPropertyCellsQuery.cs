using VisioAutomation.ShapeSheetQuery.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery.CommonQueries
{
    public class CustomPropertyCellsQuery : Query
    {
        public ColumnCellIndex SortKey { get; set; }
        public ColumnCellIndex Ask { get; set; }
        public ColumnCellIndex Calendar { get; set; }
        public ColumnCellIndex Format { get; set; }
        public ColumnCellIndex Invis { get; set; }
        public ColumnCellIndex Label { get; set; }
        public ColumnCellIndex LangID { get; set; }
        public ColumnCellIndex Prompt { get; set; }
        public ColumnCellIndex Value { get; set; }
        public ColumnCellIndex Type { get; set; }

        public CustomPropertyCellsQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionProp);


            this.SortKey = sec.AddCell(SRCCON.Prop_SortKey, nameof(SRCCON.Prop_SortKey));
            this.Ask = sec.AddCell(SRCCON.Prop_Ask, nameof(SRCCON.Prop_Ask));
            this.Calendar = sec.AddCell(SRCCON.Prop_Calendar, nameof(SRCCON.Prop_Calendar));
            this.Format = sec.AddCell(SRCCON.Prop_Format, nameof(SRCCON.Prop_Format));
            this.Invis = sec.AddCell(SRCCON.Prop_Invisible, nameof(SRCCON.Prop_Invisible));
            this.Label = sec.AddCell(SRCCON.Prop_Label, nameof(SRCCON.Prop_Label));
            this.LangID = sec.AddCell(SRCCON.Prop_LangID, nameof(SRCCON.Prop_LangID));
            this.Prompt = sec.AddCell(SRCCON.Prop_Prompt, nameof(SRCCON.Prop_Prompt));
            this.Type = sec.AddCell(SRCCON.Prop_Type, nameof(SRCCON.Prop_Type));
            this.Value = sec.AddCell(SRCCON.Prop_Value, nameof(SRCCON.Prop_Value));

        }

        public Shapes.CustomProperties.CustomPropertyCells GetCells(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Shapes.CustomProperties.CustomPropertyCells();
            cells.Value = row[this.Value];
            cells.Calendar = Extensions.CellDataMethods.ToInt(row[this.Calendar]);
            cells.Format = row[this.Format];
            cells.Invisible = Extensions.CellDataMethods.ToInt(row[this.Invis]);
            cells.Label = row[this.Label];
            cells.LangId = Extensions.CellDataMethods.ToInt(row[this.LangID]);
            cells.Prompt = row[this.Prompt];
            cells.SortKey = Extensions.CellDataMethods.ToInt(row[this.SortKey]);
            cells.Type = Extensions.CellDataMethods.ToInt(row[this.Type]);
            cells.Ask = Extensions.CellDataMethods.ToBool(row[this.Ask]);
            return cells;
        }
    }
}