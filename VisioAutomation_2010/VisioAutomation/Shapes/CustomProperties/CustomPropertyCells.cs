using VA=VisioAutomation;
using VisioAutomation.Extensions;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.CustomProperties
{
    public class CustomPropertyCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public ShapeSheet.CellData<bool> Ask { get; set; }
        public ShapeSheet.CellData<int> Calendar { get; set; }
        public ShapeSheet.CellData<double> Format { get; set; }
        public ShapeSheet.CellData<int> Invisible { get; set; }
        public ShapeSheet.CellData<double> Label { get; set; }
        public ShapeSheet.CellData<int> LangId { get; set; }
        public ShapeSheet.CellData<double> Prompt { get; set; }
        public ShapeSheet.CellData<int> SortKey { get; set; }
        public ShapeSheet.CellData<int> Type { get; set; }
        public ShapeSheet.CellData<double> Value { get; set; }

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

        public CustomPropertyCells(ShapeSheet.FormulaLiteral value)
        {
            this.Value = value;
            this.Type = 2;
        }

        private string SmartStringToFormulaString(ShapeSheet.FormulaLiteral formula, bool force_no_quoting)
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

        public override IEnumerable<SRCFormulaPair> Pairs
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

                yield return this.newpair(ShapeSheet.SRCConstants.Prop_Label, str_label);
                yield return this.newpair(ShapeSheet.SRCConstants.Prop_Value, str_value);
                yield return this.newpair(ShapeSheet.SRCConstants.Prop_Format, str_format);
                yield return this.newpair(ShapeSheet.SRCConstants.Prop_Prompt, str_prompt);
                yield return this.newpair(ShapeSheet.SRCConstants.Prop_Calendar, cp.Calendar.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Prop_LangID, cp.LangId.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Prop_SortKey, cp.SortKey.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Prop_Invisible, cp.Invisible.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Prop_Type, cp.Type.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Prop_Ask, cp.Ask.Formula);
            }
        }

        public static IList<List<CustomPropertyCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return _GetCells<CustomPropertyCells,double>(page, shapeids, query, query.GetCells);
        }

        public static IList<CustomPropertyCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells<CustomPropertyCells,double>(shape, query, query.GetCells);
        }

        private static CustomPropertyCellQuery _mCellQuery;

        private static CustomPropertyCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new CustomPropertyCellQuery();
            return _mCellQuery;
        }


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

    class CustomPropertyCellQuery : CellQuery
    {
        public CellColumn SortKey { get; set; }
        public CellColumn Ask { get; set; }
        public CellColumn Calendar { get; set; }
        public CellColumn Format { get; set; }
        public CellColumn Invis { get; set; }
        public CellColumn Label { get; set; }
        public CellColumn LangID { get; set; }
        public CellColumn Prompt { get; set; }
        public CellColumn Value { get; set; }
        public CellColumn Type { get; set; }

        public CustomPropertyCellQuery() 
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionProp);

            this.SortKey = sec.AddCell(ShapeSheet.SRCConstants.Prop_SortKey, "Prop_SortKey");
            this.Ask = sec.AddCell(ShapeSheet.SRCConstants.Prop_Ask, "Prop_Ask");
            this.Calendar = sec.AddCell(ShapeSheet.SRCConstants.Prop_Calendar, "Prop_Calendar");
            this.Format = sec.AddCell(ShapeSheet.SRCConstants.Prop_Format, "Prop_Format");
            this.Invis = sec.AddCell(ShapeSheet.SRCConstants.Prop_Invisible, "Prop_Invisible");
            this.Label = sec.AddCell(ShapeSheet.SRCConstants.Prop_Label, "Prop_Label");
            this.LangID = sec.AddCell(ShapeSheet.SRCConstants.Prop_LangID, "Prop_LangID");
            this.Prompt = sec.AddCell(ShapeSheet.SRCConstants.Prop_Prompt, "Prop_Prompt");
            this.Type = sec.AddCell(ShapeSheet.SRCConstants.Prop_Type, "Prop_Type");
            this.Value = sec.AddCell(ShapeSheet.SRCConstants.Prop_Value, "Prop_Value");

        }

        public CustomPropertyCells GetCells(IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new CustomPropertyCells();
            cells.Value = row[this.Value];
            cells.Calendar = row[this.Calendar].ToInt();
            cells.Format = row[this.Format];
            cells.Invisible = row[this.Invis].ToInt();
            cells.Label = row[this.Label];
            cells.LangId = row[this.LangID].ToInt();
            cells.Prompt = row[this.Prompt];
            cells.SortKey = row[this.SortKey].ToInt();
            cells.Type = row[this.Type].ToInt();
            cells.Ask = row[this.Ask].ToBool();
            return cells;
        }
    }

}