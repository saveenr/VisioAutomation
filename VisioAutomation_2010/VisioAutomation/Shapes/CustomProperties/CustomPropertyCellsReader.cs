using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.CustomProperties
{
    public class CustomPropertyCellsReader : MultiRowReader<Shapes.CustomProperties.CustomPropertyCells>
    {
        public SubQueryColumn SortKey { get; set; }
        public SubQueryColumn Ask { get; set; }
        public SubQueryColumn Calendar { get; set; }
        public SubQueryColumn Format { get; set; }
        public SubQueryColumn Invis { get; set; }
        public SubQueryColumn Label { get; set; }
        public SubQueryColumn LangID { get; set; }
        public SubQueryColumn Prompt { get; set; }
        public SubQueryColumn Value { get; set; }
        public SubQueryColumn Type { get; set; }

        public CustomPropertyCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionProp);


            this.SortKey = sec.AddCell(SrcConstants.Prop_SortKey, nameof(SrcConstants.Prop_SortKey));
            this.Ask = sec.AddCell(SrcConstants.Prop_Ask, nameof(SrcConstants.Prop_Ask));
            this.Calendar = sec.AddCell(SrcConstants.Prop_Calendar, nameof(SrcConstants.Prop_Calendar));
            this.Format = sec.AddCell(SrcConstants.Prop_Format, nameof(SrcConstants.Prop_Format));
            this.Invis = sec.AddCell(SrcConstants.Prop_Invisible, nameof(SrcConstants.Prop_Invisible));
            this.Label = sec.AddCell(SrcConstants.Prop_Label, nameof(SrcConstants.Prop_Label));
            this.LangID = sec.AddCell(SrcConstants.Prop_LangID, nameof(SrcConstants.Prop_LangID));
            this.Prompt = sec.AddCell(SrcConstants.Prop_Prompt, nameof(SrcConstants.Prop_Prompt));
            this.Type = sec.AddCell(SrcConstants.Prop_Type, nameof(SrcConstants.Prop_Type));
            this.Value = sec.AddCell(SrcConstants.Prop_Value, nameof(SrcConstants.Prop_Value));

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