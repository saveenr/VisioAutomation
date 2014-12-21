using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.UserDefinedCells
{
    public class UserDefinedCell : VA.ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public string Name { get; set; }
        public VA.ShapeSheet.CellData<string> Value { get; set; }
        public VA.ShapeSheet.CellData<string> Prompt { get; set; }

        public UserDefinedCell()
        {
        }

        public UserDefinedCell(string name)
        {
            UserDefinedCellsHelper.CheckValidName(name);
            this.Name = name;
        }

        public UserDefinedCell(string name, string value)
        {
            UserDefinedCellsHelper.CheckValidName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException("value");
            }

            this.Name = name;
            this.Value = value;
        }

        public UserDefinedCell(string name, string value, string prompt)
        {
            UserDefinedCellsHelper.CheckValidName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException("value");
            }
            
            this.Name = name;
            this.Value = value;
            this.Prompt = prompt;
        }

        public override IEnumerable<SRCValuePair> EnumPairs()
        {
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.User_Value, this.Value.Formula);
            yield return srcvaluepair(VA.ShapeSheet.SRCConstants.User_Prompt, this.Prompt.Formula);
        }

        public override string ToString()
        {
            string s = string.Format("(Name={0},Value={1},Prompt={2})",
                                     this.Name,
                                     this.Value,
                                     this.Prompt);
            return s;
        }

        public static IList<List<UserDefinedCell>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return _GetCells<UserDefinedCell,string>(page, shapeids, query, query.GetCells);
        }

        public static IList<UserDefinedCell> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells <UserDefinedCell,string>(shape, query, query.GetCells);
        }

        private static UserDefinedCellQuery _mCellQuery;
        private static UserDefinedCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new UserDefinedCellQuery();
            return _mCellQuery;
        }

         class UserDefinedCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.CellQuery.Column Value { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Prompt { get; set; }

            public UserDefinedCellQuery()
            {
                var sec = this.Sections.Add(IVisio.VisSectionIndices.visSectionUser);
                Value = sec.Columns.Add(VA.ShapeSheet.SRCConstants.User_Value, "Value");
                Prompt = sec.Columns.Add(VA.ShapeSheet.SRCConstants.User_Prompt, "Prompt");
            }

            public UserDefinedCell GetCells(VA.ShapeSheet.CellData<string>[] row)
            {
                var cells = new UserDefinedCell();
                cells.Value = row[Value.Ordinal];
                cells.Prompt = row[Prompt.Ordinal];
                return cells;
            }
        }
    }
}