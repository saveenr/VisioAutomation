using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.CustomProperties
{
    public class CustomPropertyCellsReader : MultiRowReader<Shapes.CustomProperties.CustomPropertyCells>
    {
        public ColumnSubQuery SortKey { get; set; }
        public ColumnSubQuery Ask { get; set; }
        public ColumnSubQuery Calendar { get; set; }
        public ColumnSubQuery Format { get; set; }
        public ColumnSubQuery Invis { get; set; }
        public ColumnSubQuery Label { get; set; }
        public ColumnSubQuery LangID { get; set; }
        public ColumnSubQuery Prompt { get; set; }
        public ColumnSubQuery Value { get; set; }
        public ColumnSubQuery Type { get; set; }

        public CustomPropertyCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionProp);


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

        public override Shapes.CustomProperties.CustomPropertyCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Shapes.CustomProperties.CustomPropertyCells();
            cells.Value = row[this.Value];
            cells.Calendar = row[this.Calendar];
            cells.Format = row[this.Format];
            cells.Invisible = row[this.Invis];
            cells.Label = row[this.Label];
            cells.LangId = row[this.LangID];
            cells.Prompt = row[this.Prompt];
            cells.SortKey = row[this.SortKey];
            cells.Type = row[this.Type];
            cells.Ask = row[this.Ask];
            return cells;
        }
    }
}