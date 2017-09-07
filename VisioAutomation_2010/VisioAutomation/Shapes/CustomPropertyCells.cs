using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class CustomPropertyCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral Ask { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Calendar { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Format { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Invisible { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Label { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LangID { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Prompt { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral SortKey { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Type { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Value { get; set; }

        public CustomPropertyCells()
        {

        }

        public CustomPropertyCells(string value)
        {
            this.Value = value;
            this.Type = 0;
        }

        public CustomPropertyCells(int value)
        {
            this.Value = value;
            this.Type = 2;
        }

        public CustomPropertyCells(double value)
        {
            this.Value = value;
            this.Type = 2;
        }

        public CustomPropertyCells(float value)
        {
            this.Value = value;
            this.Type = 2;
        }

        public CustomPropertyCells(bool value)
        {
            this.Value = value ? "TRUE" : "FALSE";
            this.Type = 3;
        }

        public CustomPropertyCells(System.DateTime value)
        {
            var current_culture = System.Globalization.CultureInfo.CurrentCulture;
            string formatted_dt = value.ToString(current_culture);
            this.Value = string.Format("DATETIME(\"{0}\")", formatted_dt);
            this.Type = 5;
        }

        public CustomPropertyCells(ShapeSheet.CellValueLiteral value)
        {
            this.Value = value;
            this.Type = 2;
        }

        private string SmartStringToFormulaString(ShapeSheet.CellValueLiteral formula, bool force_no_quoting)
        {
            if (!formula.HasValue)
            {
                return null;
            }

            if (formula.Value.Length == 0)
            {
                return VisioAutomation.Utilities.Convert.StringToFormulaString(formula.Value);
            }

            if (formula.Value[0] != '\"')
            {
                if (force_no_quoting)
                {
                    return formula.Value;
                }
                return VisioAutomation.Utilities.Convert.StringToFormulaString(formula.Value);
            }

            return formula.Value;
        }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                var cp = this;

                string str_label = this.SmartStringToFormulaString(cp.Label.Value, false);
                string str_value = null;
                if (cp.Type.Value == "0" || cp.Type.Value == null)
                {
                    // if type has no value or is a "0" then it is a string
                    str_value = this.SmartStringToFormulaString(cp.Value.Value, false);
                }
                else
                {
                    // For non-stringd don't add any extra quotes
                    str_value = this.SmartStringToFormulaString(cp.Value.Value, true);
                }
                string str_format = this.SmartStringToFormulaString(cp.Format.Value, false);
                string str_prompt = this.SmartStringToFormulaString(cp.Prompt.Value, false);

                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropLabel, str_label);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropValue, str_value);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropFormat, str_format);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropPrompt, str_prompt);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropCalendar, cp.Calendar.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropLangID, cp.LangID.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropSortKey, cp.SortKey.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropInvisible, cp.Invisible.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropType, cp.Type.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropAsk, cp.Ask.Value);
            }
        }

        public static List<List<CustomPropertyCells>> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = CustomPropertyCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids, cvt);
        }

        public static List<CustomPropertyCells> GetCells(IVisio.Shape shape, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = CustomPropertyCells.lazy_query.Value;
            return query.GetCellGroups(shape, cvt);
        }

        private static readonly System.Lazy<CustomPropertyCellsReader> lazy_query = new System.Lazy<CustomPropertyCellsReader>();

        public static CustomPropertyCells FromValue(object value)
        {
            if (value is string value_str)
            {
                return new CustomPropertyCells(value_str);
            }
            else if (value is int value_int)
            {
                return new CustomPropertyCells(value_int);
            }
            else if (value is double value_double)
            {
                return new CustomPropertyCells(value_double);
            }
            else if (value is float value_float)
            {
                return new CustomPropertyCells(value_float);
            }
            else if (value is bool value_bool)
            {
                return new CustomPropertyCells(value_bool);
            }
            else if (value is System.DateTime value_datetime)
            {
                return new CustomPropertyCells(value_datetime);
            }
            var value_type = value.GetType();
            string msg = string.Format("Unsupported type for value \"{0}\" of type \"{1}\"", value, value_type.Name);
            throw new System.ArgumentException(msg);
        }
    }
}