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

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.CustomPropLabel, this.Label);
                yield return SrcValuePair.Create(SrcConstants.CustomPropValue, this.Value);
                yield return SrcValuePair.Create(SrcConstants.CustomPropFormat, this.Format);
                yield return SrcValuePair.Create(SrcConstants.CustomPropPrompt, this.Prompt);
                yield return SrcValuePair.Create(SrcConstants.CustomPropCalendar, this.Calendar);
                yield return SrcValuePair.Create(SrcConstants.CustomPropLangID, this.LangID);
                yield return SrcValuePair.Create(SrcConstants.CustomPropSortKey, this.SortKey);
                yield return SrcValuePair.Create(SrcConstants.CustomPropInvisible, this.Invisible);
                yield return SrcValuePair.Create(SrcConstants.CustomPropType, this.Type);
                yield return SrcValuePair.Create(SrcConstants.CustomPropAsk, this.Ask);
            }
        }

        public static List<List<CustomPropertyCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(page, shapeids, type);
        }
        
        public static List<CustomPropertyCells> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(shape, type);
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

            public override CustomPropertyCells ToCellGroup(Utilities.ArraySegment<string> row)
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

        public static CustomPropertyCells Create(string value)
        {
            var cp_cells = new CustomPropertyCells();
            cp_cells.Value = value;
            cp_cells.Type = "0";
            return cp_cells;
        }

        public static CustomPropertyCells Create(int value)
        {
            var cp_cells = new CustomPropertyCells();
            cp_cells.Value = value;
            cp_cells.Type = "0";
            return cp_cells;
        }

        public static CustomPropertyCells Create(double value)
        {
            var cp_cells = new CustomPropertyCells();
            cp_cells.Value = value;
            cp_cells.Type = "2";
            return cp_cells;
        }

        public static CustomPropertyCells Create(float value)
        {
            var cp_cells = new CustomPropertyCells();
            cp_cells.Value = value;
            cp_cells.Type = "2";
            return cp_cells;
        }

        public static CustomPropertyCells Create(bool value)
        {
            var cp_cells = new CustomPropertyCells();
            cp_cells.Value = value ? "TRUE" : "FALSE";
            cp_cells.Type = "3";
            return cp_cells;
        }


        public static CustomPropertyCells Create(System.DateTime value)
        {
            var cp_cells = new CustomPropertyCells();
            var current_culture = System.Globalization.CultureInfo.CurrentCulture;
            string formatted_dt = value.ToString(current_culture);
            cp_cells.Value = string.Format("DATETIME(\"{0}\")", formatted_dt);
            cp_cells.Type = "5";
            return cp_cells;
        }


        public static CustomPropertyCells Create(CellValueLiteral value)
        {
            var cp_cells = new CustomPropertyCells();
            cp_cells.Value = value;
            cp_cells.Type = "2";
            return cp_cells;
        }

        public static CustomPropertyCells Create(object value)
        {
            if (value is string value_str)
            {
                return Create(value_str);
            }
            else if (value is int value_int)
            {
                return Create(value_int);
            }
            else if (value is double value_double)
            {
                return Create(value_double);
            }
            else if (value is float value_float)
            {
                return Create(value_float);
            }
            else if (value is bool value_bool)
            {
                return Create(value_bool);
            }
            else if (value is System.DateTime value_datetime)
            {
                return Create(value_datetime);
            }
            else if (value is CellValueLiteral value_cvl)
            {
                return Create(value_cvl);
            }

            var value_type = value.GetType();
            string msg = string.Format("Unsupported type for value \"{0}\" of type \"{1}\"", value, value_type.Name);
            throw new System.ArgumentException(msg);
        }

        public void EncodeValues()
        {
            if (this.Type.Value == null || this.Type.Value == "0")
            {
                this.Value = CellValueLiteral.EncodeValue_CP(this.Value.Value, true);
            }
            else
            {
                this.Value = CellValueLiteral.EncodeValue_CP(this.Value.Value, false);

            }
            this.Label = CellValueLiteral.EncodeValue_CP(this.Label.Value, true);
            this.Format = CellValueLiteral.EncodeValue_CP(this.Format.Value, true);
            this.Prompt = CellValueLiteral.EncodeValue_CP(this.Prompt.Value, true);
        }

    }
}