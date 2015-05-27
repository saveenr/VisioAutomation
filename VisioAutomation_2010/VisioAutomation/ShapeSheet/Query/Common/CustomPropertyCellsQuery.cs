using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheet.Query.Common
{
    class CustomPropertyCellsQuery : CellQuery
    {
        public Query.CellColumn SortKey { get; set; }
        public Query.CellColumn Ask { get; set; }
        public Query.CellColumn Calendar { get; set; }
        public Query.CellColumn Format { get; set; }
        public Query.CellColumn Invis { get; set; }
        public Query.CellColumn Label { get; set; }
        public Query.CellColumn LangID { get; set; }
        public Query.CellColumn Prompt { get; set; }
        public Query.CellColumn Value { get; set; }
        public Query.CellColumn Type { get; set; }

        public CustomPropertyCellsQuery()
        {
            var sec = this.AddSection(Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionProp);


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

        public VisioAutomation.Shapes.CustomProperties.CustomPropertyCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Shapes.CustomProperties.CustomPropertyCells();
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