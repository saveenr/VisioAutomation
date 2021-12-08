using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class CustomPropertyCells : CellRecord
    {
        public Core.CellValue Ask { get; set; }
        public Core.CellValue Calendar { get; set; }
        public Core.CellValue Format { get; set; }
        public Core.CellValue Invisible { get; set; }
        public Core.CellValue Label { get; set; }
        public Core.CellValue LangID { get; set; }
        public Core.CellValue Prompt { get; set; }
        public Core.CellValue SortKey { get; set; }
        public Core.CellValue Type { get; set; }
        public Core.CellValue Value { get; set; }

        public CustomPropertyCells()
        {
        }

        public override IEnumerable<CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.Label), Core.SrcConstants.CustomPropLabel, this.Label);
            yield return this._create(nameof(this.Value), Core.SrcConstants.CustomPropValue, this.Value);
            yield return this._create(nameof(this.Format), Core.SrcConstants.CustomPropFormat, this.Format);
            yield return this._create(nameof(this.Prompt), Core.SrcConstants.CustomPropPrompt, this.Prompt);
            yield return this._create(nameof(this.Calendar), Core.SrcConstants.CustomPropCalendar, this.Calendar);
            yield return this._create(nameof(this.LangID), Core.SrcConstants.CustomPropLangID, this.LangID);
            yield return this._create(nameof(this.SortKey), Core.SrcConstants.CustomPropSortKey, this.SortKey);
            yield return this._create(nameof(this.Invisible), Core.SrcConstants.CustomPropInvisible, this.Invisible);
            yield return this._create(nameof(this.Type), Core.SrcConstants.CustomPropType, this.Type);
            yield return this._create(nameof(this.Ask), Core.SrcConstants.CustomPropAsk, this.Ask);
        }


        public CustomPropertyCells(string value, CustomPropertyType type)
        {
            var type_int = CustomPropertyTypeToInt(type);
            this.Value = value;
            this.Type = type_int;
        }

        public CustomPropertyCells(Core.CellValue value, CustomPropertyType type)
        {
            var type_int = CustomPropertyTypeToInt(type);
            this.Value = value;
            this.Type = type_int;
        }

        public static CustomPropertyCells Create(Core.CellValue value, CustomPropertyType type)
        {
            return new CustomPropertyCells(value.Value, type);
        }

        public static int CustomPropertyTypeToInt(CustomPropertyType type)
        {
            if (type == CustomPropertyType.String)
            {
                return 0;
            }
            else if (type == CustomPropertyType.FixedList)
            {
                return 1;
            }
            else if (type == CustomPropertyType.Number)
            {
                return 2;
            }
            else if (type == CustomPropertyType.Boolean)
            {
                return 3;
            }
            else if (type == CustomPropertyType.VariableList)
            {
                return 4;
            }
            else if (type == CustomPropertyType.Date)
            {
                return 5;
            }
            else if (type == CustomPropertyType.Duration)
            {
                return 6;
            }
            else if (type == CustomPropertyType.Currency)
            {
                return 7;
            }
            else
            {
                throw new System.ArgumentOutOfRangeException(nameof(type));
            }
        }

        public CustomPropertyCells(string value)
        {
            this.Value = value;
            this.Type = CustomPropertyTypeToInt(CustomPropertyType.String);
        }

        public CustomPropertyCells(int value)
        {
            this.Value = value;
            this.Type = CustomPropertyTypeToInt(CustomPropertyType.Number);
        }

        public CustomPropertyCells(long value)
        {
            this.Value = value;
            this.Type = CustomPropertyTypeToInt(CustomPropertyType.Number);
        }

        public CustomPropertyCells(float value)
        {
            this.Value = value;
            this.Type = CustomPropertyTypeToInt(CustomPropertyType.Number);
        }

        public CustomPropertyCells(double value)
        {
            this.Value = value;
            this.Type = CustomPropertyTypeToInt(CustomPropertyType.Number);
        }

        public CustomPropertyCells(bool value)
        {
            this.Value = value;
            this.Type = CustomPropertyTypeToInt(CustomPropertyType.Boolean);
        }

        public CustomPropertyCells(Core.CellValue value)
        {
            this.Value = value;
            this.Type = CustomPropertyTypeToInt(CustomPropertyType.String);
        }

        public CustomPropertyCells(System.DateTime value)
        {
            var current_culture = System.Globalization.CultureInfo.InvariantCulture;
            string formatted_dt = value.ToString(current_culture);
            string formatted_value = string.Format("DATETIME(\"{0}\")", formatted_dt);
            this.Value = formatted_value;
            this.Type = CustomPropertyTypeToInt(CustomPropertyType.Date);
        }

        public void EncodeValues()
        {
            // only quote the value when it is a string (no type specified or type equals zero)
            bool quote = (this.Type.Value == null || this.Type.Value == "0");
            this.Value = Core.CellValue.EncodeValue(this.Value.Value, quote);
            this.Label = Core.CellValue.EncodeValue(this.Label.Value);
            this.Format = Core.CellValue.EncodeValue(this.Format.Value);
            this.Prompt = Core.CellValue.EncodeValue(this.Prompt.Value);
        }


        public static List<List<CustomPropertyCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesMultipleRows(page, shapeidpairs, type);
        }

        public static List<CustomPropertyCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeMultipleRows(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        public static CustomPropertyCells RowToRecord(VASS.Data.DataRow<string> row, VASS.Data.DataColumns cols)
        {
            var record = new CustomPropertyCells();
            var getcellvalue = queryrow_to_cellrecord(row, cols);

            record.Value = getcellvalue(nameof(Value));
            record.Calendar = getcellvalue(nameof(Calendar));
            record.Format = getcellvalue(nameof(Format));
            record.Invisible = getcellvalue(nameof(Invisible));
            record.Label = getcellvalue(nameof(Label));
            record.LangID = getcellvalue(nameof(LangID));
            record.Prompt = getcellvalue(nameof(Prompt));
            record.SortKey = getcellvalue(nameof(SortKey));
            record.Type = getcellvalue(nameof(CustomPropertyCells.Type));
            record.Ask = getcellvalue(nameof(Ask));

            return record;
        }

        public class Builder : CellRecordBuilder2<CustomPropertyCells>
        {
            public Builder() : base(CellRecordQueryType.SectionQuery, CustomPropertyCells.RowToRecord)
            {
            }
        }
    }
}