using System.Collections.Generic;
using VAQUERY=VisioAutomation.ShapeSheet.Query;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.UserDefinedCells
{
    public class UserDefinedCell : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public string Name { get; set; }
        public ShapeSheet.CellData<string> Value { get; set; }
        public ShapeSheet.CellData<string> Prompt { get; set; }

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

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.User_Value, this.Value.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.User_Prompt, this.Prompt.Formula);
            }
        }

        public override string ToString()
        {
            string s = $"(Name={this.Name},Value={this.Value},Prompt={this.Prompt})";
            return s;
        }

        public static IList<List<UserDefinedCell>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = UserDefinedCell.get_query();
            return ShapeSheet.CellGroups.CellGroupMultiRow._GetCells<UserDefinedCell, string>(page, shapeids, query, query.GetCells);
        }

        public static IList<UserDefinedCell> GetCells(IVisio.Shape shape)
        {
            var query = UserDefinedCell.get_query();
            return ShapeSheet.CellGroups.CellGroupMultiRow._GetCells<UserDefinedCell, string>(shape, query, query.GetCells);
        }

        private static UserDefinedCellQuery _mCellQuery;
        private static UserDefinedCellQuery get_query()
        {
            UserDefinedCell._mCellQuery = UserDefinedCell._mCellQuery ?? new UserDefinedCellQuery();
            return UserDefinedCell._mCellQuery;
        }

        class UserDefinedCellQuery : VAQUERY.CellQuery
        {
            public VAQUERY.CellColumn Value { get; set; }
            public VAQUERY.CellColumn Prompt { get; set; }

            public UserDefinedCellQuery()
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionUser);
                this.Value = sec.AddCell(ShapeSheet.SRCConstants.User_Value,"User");
                this.Prompt = sec.AddCell(ShapeSheet.SRCConstants.User_Prompt,"Prompt");
            }

            public UserDefinedCell GetCells(IList<ShapeSheet.CellData<string>> row)
            {
                var cells = new UserDefinedCell();
                cells.Value = row[this.Value];
                cells.Prompt = row[this.Prompt];
                return cells;
            }
        }
    }
}