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


            this.SortKey = sec.AddCell(SrcConstants.CustPropSortKey, nameof(SrcConstants.CustPropSortKey));
            this.Ask = sec.AddCell(SrcConstants.CustPropAsk, nameof(SrcConstants.CustPropAsk));
            this.Calendar = sec.AddCell(SrcConstants.CustPropCalendar, nameof(SrcConstants.CustPropCalendar));
            this.Format = sec.AddCell(SrcConstants.CustPropFormat, nameof(SrcConstants.CustPropFormat));
            this.Invis = sec.AddCell(SrcConstants.CustPropInvisible, nameof(SrcConstants.CustPropInvisible));
            this.Label = sec.AddCell(SrcConstants.CustPropLabel, nameof(SrcConstants.CustPropLabel));
            this.LangID = sec.AddCell(SrcConstants.CustPropLangId, nameof(SrcConstants.CustPropLangId));
            this.Prompt = sec.AddCell(SrcConstants.CustPropPrompt, nameof(SrcConstants.CustPropPrompt));
            this.Type = sec.AddCell(SrcConstants.CustPropType, nameof(SrcConstants.CustPropType));
            this.Value = sec.AddCell(SrcConstants.CustPropValue, nameof(SrcConstants.CustPropValue));

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