using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class CustomPropertyCellsReader : ReaderMultiRow<CustomPropertyCells>
    {
        public SectionSubQueryColumn SortKey { get; set; }
        public SectionSubQueryColumn Ask { get; set; }
        public SectionSubQueryColumn Calendar { get; set; }
        public SectionSubQueryColumn Format { get; set; }
        public SectionSubQueryColumn Invis { get; set; }
        public SectionSubQueryColumn Label { get; set; }
        public SectionSubQueryColumn LangID { get; set; }
        public SectionSubQueryColumn Prompt { get; set; }
        public SectionSubQueryColumn Value { get; set; }
        public SectionSubQueryColumn Type { get; set; }

        public CustomPropertyCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionProp);


            this.SortKey = sec.AddCell(SrcConstants.CustomPropSortKey, nameof(SrcConstants.CustomPropSortKey));
            this.Ask = sec.AddCell(SrcConstants.CustomPropAsk, nameof(SrcConstants.CustomPropAsk));
            this.Calendar = sec.AddCell(SrcConstants.CustomPropCalendar, nameof(SrcConstants.CustomPropCalendar));
            this.Format = sec.AddCell(SrcConstants.CustomPropFormat, nameof(SrcConstants.CustomPropFormat));
            this.Invis = sec.AddCell(SrcConstants.CustomPropInvisible, nameof(SrcConstants.CustomPropInvisible));
            this.Label = sec.AddCell(SrcConstants.CustomPropLabel, nameof(SrcConstants.CustomPropLabel));
            this.LangID = sec.AddCell(SrcConstants.CustomPropLangID, nameof(SrcConstants.CustomPropLangID));
            this.Prompt = sec.AddCell(SrcConstants.CustomPropPrompt, nameof(SrcConstants.CustomPropPrompt));
            this.Type = sec.AddCell(SrcConstants.CustomPropType, nameof(SrcConstants.CustomPropType));
            this.Value = sec.AddCell(SrcConstants.CustomPropValue, nameof(SrcConstants.CustomPropValue));

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