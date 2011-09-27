using VA=VisioAutomation;
using VisioAutomation.Extensions;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.CustomProperties
{
    public class CustomPropertyCells : VA.ShapeSheet.CellGroups.CellGroupForSection
    {
        public VA.ShapeSheet.CellData<double> Value{ get; set; }
        public VA.ShapeSheet.CellData<double> Prompt { get; set; }
        public VA.ShapeSheet.CellData<double> Label { get; set; }
        public VA.ShapeSheet.CellData<double> Format { get; set; }
        public VA.ShapeSheet.CellData<int> SortKey { get; set; }
        public VA.ShapeSheet.CellData<int> Invisible { get; set; }
        public VA.ShapeSheet.CellData<int> Verify { get; set; }
        public VA.ShapeSheet.CellData<int> LangId { get; set; }
        public VA.ShapeSheet.CellData<int> Calendar { get; set; }
        public VA.ShapeSheet.CellData<int> Type { get; set; }

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
        
        protected override void _Apply(ApplyFormula func, short row)
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
            return VA.ShapeSheet.CellGroups.CellGroupForSection._GetObjectsFromRowsGrouped(page, shapeids, query, get_cells_from_row);
        }

        public static IList<CustomPropertyCells> GetCells(IVisio.Shape shape)
        {
            var query = new CustomPropertyQuery();
            return VA.ShapeSheet.CellGroups.CellGroupForSection._GetObjectsFromRows(shape, query, get_cells_from_row);
        }

        private static CustomPropertyCells get_cells_from_row(CustomPropertyQuery query, VA.ShapeSheet.Data.QueryDataRow<double> row)
        {
            var cells = new CustomPropertyCells();

            cells.Value = row[query.Value];
            cells.Calendar = row[query.Calendar].ToInt();
            cells.Format = row[query.Format];
            cells.Invisible = row[query.Invis].ToInt();
            cells.Label = row[query.Label];
            cells.LangId = row[query.LangID].ToInt();
            cells.Prompt = row[query.Prompt];
            cells.SortKey = row[query.SortKey].ToInt();
            cells.Type = row[query.Type].ToInt();
            return cells;
        }
    }

    class CustomPropertyQuery : VA.ShapeSheet.Query.SectionQuery
    {
        public VA.ShapeSheet.Query.SectionQueryColumn SortKey { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Ask { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Calendar { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Format { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Invis { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Label { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn LangID { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Prompt { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Value { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Type { get; set; }

        public CustomPropertyQuery() :
            base(IVisio.VisSectionIndices.visSectionProp)
        {
            SortKey = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_SortKey, "SortKey");
            Ask = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Ask, "Ask");
            Calendar = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Calendar, "Calendar");
            Format = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Format, "Format");
            Invis = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Invisible, "Invis");
            Label = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Label, "Label");
            LangID = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_LangID, "LangID");
            Prompt = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Prompt, "Prompt");
            Type = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Value, "Type");
            Value = this.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Value, "Value");
        }
    }

}