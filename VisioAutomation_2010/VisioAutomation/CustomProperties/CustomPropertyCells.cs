using VA=VisioAutomation;
using VisioAutomation.Extensions;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.CustomProperties
{
    public class CustomPropertyCells : VA.ShapeSheet.CellGroups.CellGroupMultiRow
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

        private string SmartStringToFormulaString(VA.ShapeSheet.FormulaLiteral formula)
        {
            if (!formula.HasValue)
            {
                return null;
            }

            if (formula.Value.Length == 0)
            {
                return VA.Convert.StringToFormulaString(formula.Value);
            }

            if (formula.Value[0] != '\"')
            {
                return VA.Convert.StringToFormulaString(formula.Value);
            }

            return formula.Value;
        }

        public override void ApplyFormulasForRow(ApplyFormula func, short row)
        {
            var cp = this;

            string str_label =  this.SmartStringToFormulaString(cp.Label.Formula);
            string str_value =  this.SmartStringToFormulaString(cp.Value.Formula);
            string str_format = this.SmartStringToFormulaString(cp.Format.Formula);
            string str_prompt = this.SmartStringToFormulaString(cp.Prompt.Formula);

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
            var query = get_query();
            return _GetCells(page, shapeids, query, query.GetCells);
        }

        public static IList<CustomPropertyCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells(shape, query, query.GetCells);
        }

        private static CustomPropertyCellQuery _mCellQuery;
        private static CustomPropertyCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new CustomPropertyCellQuery();
            return _mCellQuery;
        }

        private static CustomPropertyCells get_cells_from_row(CustomPropertyCellQuery cellQuery, VA.ShapeSheet.Data.Table<VA.ShapeSheet.CellData<double>> table, int row)
        {
            var cells = new CustomPropertyCells();

            cells.Value = table[row,cellQuery.Value];
            cells.Calendar = table[row,cellQuery.Calendar].ToInt();
            cells.Format = table[row,cellQuery.Format];
            cells.Invisible = table[row,cellQuery.Invis].ToInt();
            cells.Label = table[row,cellQuery.Label];
            cells.LangId = table[row,cellQuery.LangID].ToInt();
            cells.Prompt = table[row,cellQuery.Prompt];
            cells.SortKey = table[row,cellQuery.SortKey].ToInt();
            cells.Type = table[row,cellQuery.Type].ToInt();
            return cells;
        }
    }

    class CustomPropertyCellQuery : VA.ShapeSheet.Query.CellQuery
    {
        public int SortKey { get; set; }
        public int Ask { get; set; }
        public int Calendar { get; set; }
        public int Format { get; set; }
        public int Invis { get; set; }
        public int Label { get; set; }
        public int LangID { get; set; }
        public int Prompt { get; set; }
        public int Value { get; set; }
        public int Type { get; set; }

        public CustomPropertyCellQuery() 
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionProp);

            SortKey = sec.AddColumn(VA.ShapeSheet.SRCConstants.Prop_SortKey, "SortKey");
            Ask = sec.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Ask, "Ask");
            Calendar = sec.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Calendar, "Calendar");
            Format = sec.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Format, "Format");
            Invis = sec.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Invisible, "Invis");
            Label = sec.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Label, "Label");
            LangID = sec.AddColumn(VA.ShapeSheet.SRCConstants.Prop_LangID, "LangID");
            Prompt = sec.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Prompt, "Prompt");
            Type = sec.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Type, "Type");
            Value = sec.AddColumn(VA.ShapeSheet.SRCConstants.Prop_Value, "Value");
        }

        public CustomPropertyCells GetCells(VA.ShapeSheet.CellData<double>[] row)
        {
            var cells = new CustomPropertyCells();
            cells.Value = row[Value];
            cells.Calendar = row[Calendar].ToInt();
            cells.Format = row[Format];
            cells.Invisible = row[Invis].ToInt();
            cells.Label = row[Label];
            cells.LangId = row[LangID].ToInt();
            cells.Prompt = row[Prompt];
            cells.SortKey = row[SortKey].ToInt();
            cells.Type = row[Type].ToInt();
            return cells;
        }
    }

}