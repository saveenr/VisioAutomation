using System;
using System.Collections.Generic;
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
                throw new System.ArgumentNullException(nameof(value));
            }

            this.Name = name;
            this.Value = value;
        }

        public UserDefinedCell(string name, string value, string prompt)
        {
            UserDefinedCellsHelper.CheckValidName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException(nameof(value));
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
            string s = string.Format("(Name={0},Value={1},Prompt={2})", this.Name, this.Value, this.Prompt);
            return s;
        }

        public static IList<List<UserDefinedCell>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = UserDefinedCell.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroupMultiRow._GetCells<UserDefinedCell, string>(page, shapeids, query, query.GetCells);
        }

        public static IList<UserDefinedCell> GetCells(IVisio.Shape shape)
        {
            var query = UserDefinedCell.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroupMultiRow._GetCells<UserDefinedCell, string>(shape, query, query.GetCells);
        }

        private static System.Lazy<ShapeSheet.Query.Common.UserDefinedCellsQuery> lazy_query = new System.Lazy<ShapeSheet.Query.Common.UserDefinedCellsQuery>();


    }
}