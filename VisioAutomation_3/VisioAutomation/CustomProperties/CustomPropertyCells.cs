using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.CustomProperties
{
    public class CustomPropertyCells
    {
        public VA.ShapeSheet.CellData<double> Value{ get; set; }
        public VA.ShapeSheet.CellData<double> Prompt { get; set; }
        public VA.ShapeSheet.CellData<double> Label { get; set; }
        public VA.ShapeSheet.CellData<double> Format { get; set; }
        public VA.ShapeSheet.CellData<int> SortKey { get; set; }
        public VA.ShapeSheet.CellData<int> Invisible { get; set; }
        public VA.ShapeSheet.CellData<int> Verify { get; set; }
        public VA.ShapeSheet.CellData<int> LangId { get; set; }
        public VA.ShapeSheet.CellData<Calendar> Calendar { get; set; }
        public VA.ShapeSheet.CellData<Format> Type { get; set; }

        internal static readonly VA.CustomProperties.CustomPropertyQuery custprop_query = new VA.CustomProperties.CustomPropertyQuery();

        public CustomPropertyCells()
        {
            
        }

        public CustomPropertyCells(VA.ShapeSheet.FormulaLiteral value)
        {
            this.Value = value;
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update, short row)
        {
            var cp = this;

            string str_label = cp.Label.Formula.HasValue ? cp.Label.Formula.Encode() : null;
            string str_value = cp.Value.Formula.HasValue ? cp.Value.Formula.Encode() : null;
            string str_format = cp.Format.Formula.HasValue ? cp.Format.Formula.Encode() : null;
            string str_prompt = cp.Prompt.Formula.HasValue ? cp.Prompt.Formula.Encode() : null;

            update.SetFormulaIgnoreNull(CustomPropertyCells.custprop_query.GetCellSRCForRow(CustomPropertyCells.custprop_query.Label, row), str_label);
            update.SetFormulaIgnoreNull(CustomPropertyCells.custprop_query.GetCellSRCForRow(CustomPropertyCells.custprop_query.Value, row), str_value);
            update.SetFormulaIgnoreNull(CustomPropertyCells.custprop_query.GetCellSRCForRow(CustomPropertyCells.custprop_query.Format, row), str_format);
            update.SetFormulaIgnoreNull(CustomPropertyCells.custprop_query.GetCellSRCForRow(CustomPropertyCells.custprop_query.Prompt, row), str_prompt);
            update.SetFormulaIgnoreNull(CustomPropertyCells.custprop_query.GetCellSRCForRow(CustomPropertyCells.custprop_query.Calendar, row), cp.Calendar.Formula);
            update.SetFormulaIgnoreNull(CustomPropertyCells.custprop_query.GetCellSRCForRow(CustomPropertyCells.custprop_query.LangID, row), cp.LangId.Formula);
            update.SetFormulaIgnoreNull(CustomPropertyCells.custprop_query.GetCellSRCForRow(CustomPropertyCells.custprop_query.SortKey, row), cp.SortKey.Formula);
            update.SetFormulaIgnoreNull(CustomPropertyCells.custprop_query.GetCellSRCForRow(CustomPropertyCells.custprop_query.Invis, row), cp.Invisible.Formula);
            update.SetFormulaIgnoreNull(CustomPropertyCells.custprop_query.GetCellSRCForRow(CustomPropertyCells.custprop_query.Type, row), cp.Type.Formula);
        }

    }
}