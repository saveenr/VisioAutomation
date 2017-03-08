using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.CustomProperties
{
    public class CustomPropertyCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public ShapeSheet.CellData Ask { get; set; }
        public ShapeSheet.CellData Calendar { get; set; }
        public ShapeSheet.CellData Format { get; set; }
        public ShapeSheet.CellData Invisible { get; set; }
        public ShapeSheet.CellData Label { get; set; }
        public ShapeSheet.CellData LangID { get; set; }
        public ShapeSheet.CellData Prompt { get; set; }
        public ShapeSheet.CellData SortKey { get; set; }
        public ShapeSheet.CellData Type { get; set; }
        public ShapeSheet.CellData Value { get; set; }

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
                return formula.Encode();
            }

            if (formula.Value[0] != '\"')
            {
                if (force_no_quoting)
                {
                    return formula.Value;
                }
                return formula.Encode();
            }

            return formula.Value;
        }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                var cp = this;

                string str_label = this.SmartStringToFormulaString(cp.Label.Formula, false);
                string str_value = null;
                if (cp.Type.Formula.Value == "0" || cp.Type.Formula.Value == null)
                {
                    // if type has no value or is a "0" then it is a string
                    str_value = this.SmartStringToFormulaString(cp.Value.Formula, false);
                }
                else
                {
                    // For non-stringd don't add any extra quotes
                    str_value = this.SmartStringToFormulaString(cp.Value.Formula, true);
                }
                string str_format = this.SmartStringToFormulaString(cp.Format.Formula, false);
                string str_prompt = this.SmartStringToFormulaString(cp.Prompt.Formula, false);

                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropLabel, str_label);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropValue, str_value);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropFormat, str_format);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropPrompt, str_prompt);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropCalendar, cp.Calendar.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropLangID, cp.LangID.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropSortKey, cp.SortKey.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropInvisible, cp.Invisible.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropType, cp.Type.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.CustomPropAsk, cp.Ask.Formula);
            }
        }

        public static List<List<CustomPropertyCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = CustomPropertyCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static List<CustomPropertyCells> GetCells(IVisio.Shape shape)
        {
            var query = CustomPropertyCells.lazy_query.Value;
            return query.GetCellGroups(shape);
        }

        private static readonly System.Lazy<CustomPropertyCellsReader> lazy_query = new System.Lazy<CustomPropertyCellsReader>();

        public static CustomPropertyCells FromValue(object value)
        {
            if (value is string)
            {
                return new CustomPropertyCells((string) value);
            }
            else if (value is int)
            {
                return new CustomPropertyCells((int) value);
            }
            else if (value is double)
            {
                return new CustomPropertyCells((double) value);
            }
            else if (value is float)
            {
                return new CustomPropertyCells((float) value);
            }
            else if (value is bool)
            {
                return new CustomPropertyCells((bool) value);
            }
            else if (value is System.DateTime)
            {
                return new CustomPropertyCells((System.DateTime) value);
            }
            else
            {
                string msg = string.Format("Unsupported type for value \"{0}\" \"{1}\"", value, value.GetType());
                throw new System.ArgumentOutOfRangeException(msg);
            }
        }
    }
}