using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class CustomPropertyCells : CellGroupMultiRow
    {
        public CellValueLiteral Ask { get; set; }
        public CellValueLiteral Calendar { get; set; }
        public CellValueLiteral Format { get; set; }
        public CellValueLiteral Invisible { get; set; }
        public CellValueLiteral Label { get; set; }
        public CellValueLiteral LangID { get; set; }
        public CellValueLiteral Prompt { get; set; }
        public CellValueLiteral SortKey { get; set; }
        public CellValueLiteral Type { get; set; }
        public CellValueLiteral Value { get; set; }

        public CustomPropertyCells()
        {

        }

        private string SmartStringToFormulaString(CellValueLiteral formula, bool force_no_quoting)
        {
            if (!formula.HasValue)
            {
                return null;
            }

            if (formula.Value.Length == 0)
            {
                return Utilities.Convert.StringToFormulaString(formula.Value);
            }

            if (formula.Value[0] != '\"')
            {
                if (force_no_quoting)
                {
                    return formula.Value;
                }
                return Utilities.Convert.StringToFormulaString(formula.Value);
            }

            return formula.Value;
        }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                // Label
                string str_label = this.SmartStringToFormulaString(this.Label.Value, false);

                // Value
                string str_format = this.SmartStringToFormulaString(this.Format.Value, false);

                // Prompt
                string str_prompt = this.SmartStringToFormulaString(this.Prompt.Value, false);

                // Value
                string str_value = null;
                if (this.Type.Value == "0" || this.Type.Value == null)
                {
                    // if type has no value or is a "0" then it is a string
                    str_value = this.SmartStringToFormulaString(this.Value.Value, false);
                }
                else
                {
                    // For non-stringd don't add any extra quotes
                    str_value = this.SmartStringToFormulaString(this.Value.Value, true);
                }

                yield return SrcValuePair.Create(SrcConstants.CustomPropLabel, str_label);
                yield return SrcValuePair.Create(SrcConstants.CustomPropValue, str_value);
                yield return SrcValuePair.Create(SrcConstants.CustomPropFormat, str_format);
                yield return SrcValuePair.Create(SrcConstants.CustomPropPrompt, str_prompt);
                yield return SrcValuePair.Create(SrcConstants.CustomPropCalendar, this.Calendar);
                yield return SrcValuePair.Create(SrcConstants.CustomPropLangID, this.LangID);
                yield return SrcValuePair.Create(SrcConstants.CustomPropSortKey, this.SortKey);
                yield return SrcValuePair.Create(SrcConstants.CustomPropInvisible, this.Invisible);
                yield return SrcValuePair.Create(SrcConstants.CustomPropType, this.Type);
                yield return SrcValuePair.Create(SrcConstants.CustomPropAsk, this.Ask);
            }
        }

        public static List<List<CustomPropertyCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(page, shapeids, cvt);
        }
        
        public static List<CustomPropertyCells> GetCells(IVisio.Shape shape, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<CustomPropertyCellsReader> lazy_query = new System.Lazy<CustomPropertyCellsReader>();


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
                var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionProp);


                this.SortKey = sec.Columns.Add(SrcConstants.CustomPropSortKey, nameof(this.SortKey));
                this.Ask = sec.Columns.Add(SrcConstants.CustomPropAsk, nameof(this.Ask));
                this.Calendar = sec.Columns.Add(SrcConstants.CustomPropCalendar, nameof(this.Calendar));
                this.Format = sec.Columns.Add(SrcConstants.CustomPropFormat, nameof(this.Format));
                this.Invis = sec.Columns.Add(SrcConstants.CustomPropInvisible, nameof(this.Invis));
                this.Label = sec.Columns.Add(SrcConstants.CustomPropLabel, nameof(this.Label));
                this.LangID = sec.Columns.Add(SrcConstants.CustomPropLangID, nameof(this.LangID));
                this.Prompt = sec.Columns.Add(SrcConstants.CustomPropPrompt, nameof(this.Prompt));
                this.Type = sec.Columns.Add(SrcConstants.CustomPropType, nameof(this.Type));
                this.Value = sec.Columns.Add(SrcConstants.CustomPropValue, nameof(this.Value));

            }

            public override CustomPropertyCells CellDataToCellGroup(Utilities.ArraySegment<string> row)
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
}