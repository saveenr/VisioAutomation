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


            this.SortKey = sec.AddCell(ShapeSheet.SRCConstants.Prop_SortKey, nameof(ShapeSheet.SRCConstants.Prop_SortKey));
            this.Ask = sec.AddCell(ShapeSheet.SRCConstants.Prop_Ask, nameof(ShapeSheet.SRCConstants.Prop_Ask));
            this.Calendar = sec.AddCell(ShapeSheet.SRCConstants.Prop_Calendar, nameof(ShapeSheet.SRCConstants.Prop_Calendar));
            this.Format = sec.AddCell(ShapeSheet.SRCConstants.Prop_Format, nameof(ShapeSheet.SRCConstants.Prop_Format));
            this.Invis = sec.AddCell(ShapeSheet.SRCConstants.Prop_Invisible, nameof(ShapeSheet.SRCConstants.Prop_Invisible));
            this.Label = sec.AddCell(ShapeSheet.SRCConstants.Prop_Label, nameof(ShapeSheet.SRCConstants.Prop_Label));
            this.LangID = sec.AddCell(ShapeSheet.SRCConstants.Prop_LangID, nameof(ShapeSheet.SRCConstants.Prop_LangID));
            this.Prompt = sec.AddCell(ShapeSheet.SRCConstants.Prop_Prompt, nameof(ShapeSheet.SRCConstants.Prop_Prompt));
            this.Type = sec.AddCell(ShapeSheet.SRCConstants.Prop_Type, nameof(ShapeSheet.SRCConstants.Prop_Type));
            this.Value = sec.AddCell(ShapeSheet.SRCConstants.Prop_Value, nameof(ShapeSheet.SRCConstants.Prop_Value));

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