using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public class CustomPropertyCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValueLiteral Ask { get; set; }
        public VASS.CellValueLiteral Calendar { get; set; }
        public VASS.CellValueLiteral Format { get; set; }
        public VASS.CellValueLiteral Invisible { get; set; }
        public VASS.CellValueLiteral Label { get; set; }
        public VASS.CellValueLiteral LangID { get; set; }
        public VASS.CellValueLiteral Prompt { get; set; }
        public VASS.CellValueLiteral SortKey { get; set; }
        public VASS.CellValueLiteral Type { get; set; }
        public VASS.CellValueLiteral Value { get; set; }

        public CustomPropertyCells()
        {

        }

        public override IEnumerable<VASS.CellGroups.CellMetadataItem> CellMetadata
        {
            get
            {
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Label), VASS.SrcConstants.CustomPropLabel, this.Label);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Value), VASS.SrcConstants.CustomPropValue, this.Value);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Format), VASS.SrcConstants.CustomPropFormat, this.Format);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Prompt), VASS.SrcConstants.CustomPropPrompt, this.Prompt);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Calendar), VASS.SrcConstants.CustomPropCalendar, this.Calendar);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.LangID), VASS.SrcConstants.CustomPropLangID, this.LangID);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.SortKey), VASS.SrcConstants.CustomPropSortKey, this.SortKey);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Invisible), VASS.SrcConstants.CustomPropInvisible, this.Invisible);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Type), VASS.SrcConstants.CustomPropType, this.Type);
                yield return VASS.CellGroups.CellMetadataItem.Create(nameof(this.Ask), VASS.SrcConstants.CustomPropAsk, this.Ask);
            }
        }


        public CustomPropertyCells(string value, CustomPropertyType type)
        {
            var type_int = CustomPropertyTypeToInt(type);
            this.Value = value;
            this.Type = type_int;
        }

        public CustomPropertyCells(VASS.CellValueLiteral value, CustomPropertyType type)
        {
            var type_int = CustomPropertyTypeToInt(type);
            this.Value = value;
            this.Type = type_int;
        }
        
        public static CustomPropertyCells Create(VASS.CellValueLiteral value, CustomPropertyType type)
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

        public CustomPropertyCells(VASS.CellValueLiteral value)
        {
            this.Value = value;
            this.Type = CustomPropertyTypeToInt(CustomPropertyType.String);
        }

        public CustomPropertyCells(System.DateTime value)
        {
            var current_culture = System.Globalization.CultureInfo.InvariantCulture;
            string formatted_dt = value.ToString(current_culture);
            string _Value = string.Format("DATETIME(\"{0}\")", formatted_dt);
            this.Value = _Value;
            this.Type = CustomPropertyTypeToInt(CustomPropertyType.Date);
        }

        public void EncodeValues()
        {
            // only quote the value when it is a string (no type specified or type equals zero)
            bool quote = (this.Type.Value == null || this.Type.Value == "0");
            this.Value = VASS.CellValueLiteral.EncodeValue(this.Value.Value, quote);
            this.Label = VASS.CellValueLiteral.EncodeValue(this.Label.Value);
            this.Format = VASS.CellValueLiteral.EncodeValue(this.Format.Value);
            this.Prompt = VASS.CellValueLiteral.EncodeValue(this.Prompt.Value);
        }

        private void Validate()
        {
            if (!this.Prompt.ValidateValue(true))
            {
                throw new System.ArgumentException("Invalid value for Custom Property's Prompt");
            }

            if (!this.Label.ValidateValue(true))
            {
                throw new System.ArgumentException("Invalid value for Custom Property's Label");
            }

            if (!this.Format.ValidateValue(true))
            {
                throw new System.ArgumentException("Invalid value for Custom Property's Format");
            }

            if (!this.Value.ValidateValue(false))
            {
                //throw new System.ArgumentException("Invalid value for Custom Property's Value");
            }
        }
    }
}