using VA=VisioAutomation;
using VisioAutomation.Extensions;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.CustomProperties
{
    public class CustomPropertyCells : VA.ShapeSheet.CellSectionDataGroup
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

        private string encode_if_needed(VA.ShapeSheet.FormulaLiteral f)
        {
            if (!f.HasValue)
            {
                return null;
            }

            if (f.Value.Length==0)
            {
                return VA.Convert.StringToFormulaString(f.Value);
            }

            if (f.Value[0]!='\"')
            {
                return VA.Convert.StringToFormulaString(f.Value);                
            }

            return f.Value;
        }
        
        protected override void _Apply(VA.ShapeSheet.CellSectionDataGroup.ApplyFormula func, short row)
        {
            var cp = this;

            string str_label = encode_if_needed(cp.Label.Formula);
            string str_value = encode_if_needed(cp.Value.Formula);
            string str_format = encode_if_needed(cp.Format.Formula);
            string str_prompt = encode_if_needed(cp.Prompt.Formula);

            func(VA.ShapeSheet.SRCConstants.Prop_Label.ForRow(row), str_label);
            func(VA.ShapeSheet.SRCConstants.Prop_Value.ForRow( row), str_value);
            func(VA.ShapeSheet.SRCConstants.Prop_Format.ForRow( row), str_format);
            func(VA.ShapeSheet.SRCConstants.Prop_Prompt.ForRow( row), str_prompt);
            func(VA.ShapeSheet.SRCConstants.Prop_Calendar.ForRow( row), cp.Calendar.Formula);
            func(VA.ShapeSheet.SRCConstants.Prop_LangID.ForRow( row), cp.LangId.Formula);
            func(VA.ShapeSheet.SRCConstants.Prop_SortKey.ForRow( row), cp.SortKey.Formula);
            func(VA.ShapeSheet.SRCConstants.Prop_Invisible.ForRow( row), cp.Invisible.Formula);
            func(VA.ShapeSheet.SRCConstants.Prop_Type.ForRow( row), cp.Type.Formula);
        }

        public static IList<List<CustomPropertyCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new CustomPropertyQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetCells(page, shapeids, query, get_cells_from_row);
        }

        public static IList<CustomPropertyCells> GetCells(IVisio.Shape shape)
        {
            var query = new CustomPropertyQuery();
            return VA.ShapeSheet.CellSectionDataGroup._GetCells(shape, query, get_cells_from_row);
        }

        private static CustomPropertyCells get_cells_from_row(CustomPropertyQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new CustomPropertyCells();

            cells.Value = qds.GetItem(row, CustomPropertyCells.custprop_query.Value);
            cells.Calendar = qds.GetItem(row, CustomPropertyCells.custprop_query.Calendar, v => (VA.CustomProperties.Calendar)v);
            cells.Format = qds.GetItem(row, CustomPropertyCells.custprop_query.Format);
            cells.Invisible = qds.GetItem(row, CustomPropertyCells.custprop_query.Invis, v => (int)v);
            cells.Label = qds.GetItem(row, CustomPropertyCells.custprop_query.Label);
            cells.LangId = qds.GetItem(row, CustomPropertyCells.custprop_query.LangID, v => (int)v);
            cells.Prompt = qds.GetItem(row, CustomPropertyCells.custprop_query.Prompt);
            cells.SortKey = qds.GetItem(row, CustomPropertyCells.custprop_query.SortKey, v => (int)v);
            cells.Type = qds.GetItem(row, CustomPropertyCells.custprop_query.Type, v => (VA.CustomProperties.Format)((int)v));
            return cells;
        }
    }
}