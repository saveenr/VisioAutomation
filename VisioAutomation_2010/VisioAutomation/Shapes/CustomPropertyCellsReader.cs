using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class CustomPropertyCellsReader : ReaderMultiRow<CustomPropertyCells>
    {
        public SectionQueryColumn SortKey { get; set; }
        public SectionQueryColumn Ask { get; set; }
        public SectionQueryColumn Calendar { get; set; }
        public SectionQueryColumn Format { get; set; }
        public SectionQueryColumn Invis { get; set; }
        public SectionQueryColumn Label { get; set; }
        public SectionQueryColumn LangID { get; set; }
        public SectionQueryColumn Prompt { get; set; }
        public SectionQueryColumn Value { get; set; }
        public SectionQueryColumn Type { get; set; }

        public CustomPropertyCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionProp);


            this.SortKey = sec.Columns.Add(SrcConstants.CustomPropSortKey, nameof(SrcConstants.CustomPropSortKey));
            this.Ask = sec.Columns.Add(SrcConstants.CustomPropAsk, nameof(SrcConstants.CustomPropAsk));
            this.Calendar = sec.Columns.Add(SrcConstants.CustomPropCalendar, nameof(SrcConstants.CustomPropCalendar));
            this.Format = sec.Columns.Add(SrcConstants.CustomPropFormat, nameof(SrcConstants.CustomPropFormat));
            this.Invis = sec.Columns.Add(SrcConstants.CustomPropInvisible, nameof(SrcConstants.CustomPropInvisible));
            this.Label = sec.Columns.Add(SrcConstants.CustomPropLabel, nameof(SrcConstants.CustomPropLabel));
            this.LangID = sec.Columns.Add(SrcConstants.CustomPropLangID, nameof(SrcConstants.CustomPropLangID));
            this.Prompt = sec.Columns.Add(SrcConstants.CustomPropPrompt, nameof(SrcConstants.CustomPropPrompt));
            this.Type = sec.Columns.Add(SrcConstants.CustomPropType, nameof(SrcConstants.CustomPropType));
            this.Value = sec.Columns.Add(SrcConstants.CustomPropValue, nameof(SrcConstants.CustomPropValue));

        }

        public override CustomPropertyCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new CustomPropertyCells();
            cells.Value = row[this.Value];
            cells.Calendar = row[this.Calendar];
            cells.Format = row[this.Format];
            cells.Invisible = row[this.Invis];
            cells.Label = row[this.Label];
            cells.LangID = row[this.LangID];
            cells.Prompt = row[this.Prompt];
            cells.SortKey = row[this.SortKey];
            cells.Type = row[this.Type];
            cells.Ask = row[this.Ask];
            return cells;
        }
    }
}