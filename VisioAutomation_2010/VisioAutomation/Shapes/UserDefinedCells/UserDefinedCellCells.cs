using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.UserDefinedCells
{
    public class UserDefinedCellCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public string Name { get; set; }
        public ShapeSheet.CellData Value { get; set; }
        public ShapeSheet.CellData Prompt { get; set; }

        public UserDefinedCellCells()
        {
        }

        public UserDefinedCellCells(string name)
        {
            UserDefinedCellHelper.CheckValidName(name);
            this.Name = name;
        }

        public UserDefinedCellCells(string name, string value)
        {
            UserDefinedCellHelper.CheckValidName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException(nameof(value));
            }

            this.Name = name;
            this.Value = value;
        }

        public UserDefinedCellCells(string name, string value, string prompt)
        {
            UserDefinedCellHelper.CheckValidName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException(nameof(value));
            }
            
            this.Name = name;
            this.Value = value;
            this.Prompt = prompt;
        }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.UserDelCellValue, this.Value.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.UserDefCellPrompt, this.Prompt.Formula);
            }
        }

        public override string ToString()
        {
            string s = string.Format("(Name={0},Value={1},Prompt={2})", this.Name, this.Value, this.Prompt);
            return s;
        }

        public static List<List<UserDefinedCellCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = UserDefinedCellCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static List<UserDefinedCellCells> GetCells(IVisio.Shape shape)
        {
            var query = UserDefinedCellCells.lazy_query.Value;
            return query.GetCellGroups(shape);
        }

        private static readonly System.Lazy<UserDefinedCellCellsReader> lazy_query = new System.Lazy<UserDefinedCellCellsReader>();
    }
}