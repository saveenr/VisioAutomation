﻿using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class CustomPropertyCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValue Ask { get; set; }
        public VASS.CellValue Calendar { get; set; }
        public VASS.CellValue Format { get; set; }
        public VASS.CellValue Invisible { get; set; }
        public VASS.CellValue Label { get; set; }
        public VASS.CellValue LangID { get; set; }
        public VASS.CellValue Prompt { get; set; }
        public VASS.CellValue SortKey { get; set; }
        public VASS.CellValue Type { get; set; }
        public VASS.CellValue Value { get; set; }

        public CustomPropertyCells()
        {

        }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.Label), VASS.SrcConstants.CustomPropLabel, this.Label);
            yield return this.Create(nameof(this.Value), VASS.SrcConstants.CustomPropValue, this.Value);
            yield return this.Create(nameof(this.Format), VASS.SrcConstants.CustomPropFormat, this.Format);
            yield return this.Create(nameof(this.Prompt), VASS.SrcConstants.CustomPropPrompt, this.Prompt);
            yield return this.Create(nameof(this.Calendar), VASS.SrcConstants.CustomPropCalendar, this.Calendar);
            yield return this.Create(nameof(this.LangID), VASS.SrcConstants.CustomPropLangID, this.LangID);
            yield return this.Create(nameof(this.SortKey), VASS.SrcConstants.CustomPropSortKey, this.SortKey);
            yield return this.Create(nameof(this.Invisible), VASS.SrcConstants.CustomPropInvisible, this.Invisible);
            yield return this.Create(nameof(this.Type), VASS.SrcConstants.CustomPropType, this.Type);
            yield return this.Create(nameof(this.Ask), VASS.SrcConstants.CustomPropAsk, this.Ask);
        }


        public CustomPropertyCells(string value, CustomPropertyType type)
        {
            var type_int = CustomPropertyTypeToInt(type);
            this.Value = value;
            this.Type = type_int;
        }

        public CustomPropertyCells(VASS.CellValue value, CustomPropertyType type)
        {
            var type_int = CustomPropertyTypeToInt(type);
            this.Value = value;
            this.Type = type_int;
        }
        
        public static CustomPropertyCells Create(VASS.CellValue value, CustomPropertyType type)
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

        public CustomPropertyCells(VASS.CellValue value)
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
            this.Value = VASS.CellValue.EncodeValue(this.Value.Value, quote);
            this.Label = VASS.CellValue.EncodeValue(this.Label.Value);
            this.Format = VASS.CellValue.EncodeValue(this.Format.Value);
            this.Prompt = VASS.CellValue.EncodeValue(this.Prompt.Value);
        }


        public static List<List<CustomPropertyCells>> GetCells(IVisio.Page page, ShapeIDPairs shapeidpairs, VASS.CellValueType type)
        {
            var reader = Custom_Property_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeidpairs, type);
        }

        public static List<CustomPropertyCells> GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = Custom_Property_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<CustomPropertyCellsBuilder> Custom_Property_lazy_builder = new System.Lazy<CustomPropertyCellsBuilder>();


        public class CustomPropertyCellsBuilder : VASS.CellGroups.CellGroupBuilder<CustomPropertyCells>
        {

            public CustomPropertyCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.MultiRow)
            {
            }

            public override CustomPropertyCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new CustomPropertyCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                cells.Value = getcellvalue(nameof(CustomPropertyCells.Value));
                cells.Calendar = getcellvalue(nameof(CustomPropertyCells.Calendar));
                cells.Format = getcellvalue(nameof(CustomPropertyCells.Format));
                cells.Invisible = getcellvalue(nameof(CustomPropertyCells.Invisible));
                cells.Label = getcellvalue(nameof(CustomPropertyCells.Label));
                cells.LangID = getcellvalue(nameof(CustomPropertyCells.LangID));
                cells.Prompt = getcellvalue(nameof(CustomPropertyCells.Prompt));
                cells.SortKey = getcellvalue(nameof(CustomPropertyCells.SortKey));
                cells.Type = getcellvalue(nameof(CustomPropertyCells.Type));
                cells.Ask = getcellvalue(nameof(CustomPropertyCells.Ask));

                return cells;
            }
        }

    }
}