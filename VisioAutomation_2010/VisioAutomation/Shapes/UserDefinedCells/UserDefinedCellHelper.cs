using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Update;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.UserDefinedCells
{
    public static class UserDefinedCellHelper
    {
        private static readonly short _userdefinedcell_section = ShapeSheet.SRCConstants.User_Prompt.Section;

        private static string GetRowName(string name)
        {
            return "User." + name;
        }

        public static void Delete(IVisio.Shape shape, string name)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            UserDefinedCellHelper.CheckValidName(name);

            string full_prop_name = UserDefinedCellHelper.GetRowName(name);

            short row = shape.CellsU[full_prop_name].Row;
            shape.DeleteRow(UserDefinedCellHelper._userdefinedcell_section, row);
        }

        public static void Set(IVisio.Shape shape, string name, ShapeSheet.CellData<double> value, ShapeSheet.CellData<double> prompt)
        {
            UserDefinedCellHelper.Set(shape, name, value.Formula.Value, prompt.Formula.Value);
        }

        public static void Set(IVisio.Shape shape, string name, ShapeSheet.FormulaLiteral value, ShapeSheet.FormulaLiteral prompt)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            UserDefinedCellHelper.CheckValidName(name);

            if (UserDefinedCellHelper.Contains(shape, name))
            {
                string full_prop_name = UserDefinedCellHelper.GetRowName(name);

                if (value.HasValue)
                {
                    string value_cell_name = full_prop_name;
                    var cell = shape.CellsU[value_cell_name];
                    cell.FormulaU = value.Encode();                    
                }

                if (prompt.HasValue)
                {
                    string prompt_cell_name = full_prop_name+".Prompt";
                    var cell = shape.CellsU[prompt_cell_name];
                    cell.FormulaU = prompt.Encode();                                        
                }
                return;
            }

            short row = shape.AddNamedRow(
                UserDefinedCellHelper._userdefinedcell_section,
                name,
                (short)IVisio.VisRowIndices.visRowUser);

            var update = new Update();

            if (value.HasValue)
            {
                var src = new ShapeSheet.SRC(UserDefinedCellHelper._userdefinedcell_section, row, (short)IVisio.VisCellIndices.visUserValue);
                update.SetFormula(src, value.Encode());
            }

            if (prompt.HasValue)
            {
                var src = new ShapeSheet.SRC(UserDefinedCellHelper._userdefinedcell_section, row, (short)IVisio.VisCellIndices.visUserPrompt);
                update.SetFormula(src, prompt.Encode());
            }

            update.Execute(shape);
        }

        /// <summary>
        /// Gets all the user properties defined on a shape
        /// </summary>
        /// <remarks>
        /// If there are no user properties then null will be returned</remarks>
        /// <param name="shape"></param>
        /// <returns>A list of user  properties</returns>
        public static IList<UserDefinedCell> Get(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var prop_count = UserDefinedCellHelper.GetCount(shape);
            if (prop_count < 1)
            {
                return new List<UserDefinedCell>(0);
            }

            var prop_names = UserDefinedCellHelper.GetNames(shape);
            if (prop_names.Count != prop_count)
            {
                throw new AutomationException("Unexpected number of prop names");
            }

            var shape_data = UserDefinedCell.GetCells(shape);

            var list = new List<UserDefinedCell>(prop_count);
            for (int i = 0; i < prop_count; i++)
            {
                shape_data[i].Name = prop_names[i];
                list.Add(shape_data[i]);
            }

            return list;
        }

        public static IList<List<UserDefinedCell>> Get(IVisio.Page page, IList<IVisio.Shape> shapes)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            var shapeids = shapes.Select(s => s.ID).ToList();

            var list_data = UserDefinedCell.GetCells(page,shapeids);

            var list_list = new List<List<UserDefinedCell>>(shapeids.Count);

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var shape_data = list_data[i];
                var prop_names = UserDefinedCellHelper.GetNames(shape);

                var list = new List<UserDefinedCell>(shape_data.Count);
                list_list.Add(list);
                for (int j = 0; j < shape_data.Count ; j++)
                {
                    shape_data[j].Name = prop_names[j];
                    list.Add(shape_data[j]);
                }
            }

            return list_list;
        }

        /// <summary>
        /// Get the number of user-defined cells for the shape.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            // If the User Property section does not exist then return zero immediately
            if (0 == shape.SectionExists[UserDefinedCellHelper._userdefinedcell_section, (short)IVisio.VisExistsFlags.visExistsAnywhere])
            {
                return 0;
            }

            var section = shape.Section[UserDefinedCellHelper._userdefinedcell_section];

            if (section == null)
            {
                string msg = string.Format("Could not find the user-defined section for shape {0}", shape.NameU);
                throw new AutomationException(msg);
            }

            int row_count = section.Shape.RowCount[UserDefinedCellHelper._userdefinedcell_section];

            return row_count;
        }

        /// <summary>
        /// Returns all the Names of the user-defined cells
        /// </summary>
        /// <remarks>
        /// names of user defined cells are not queryable get GetResults & GetFormulas
        /// </remarks>
        /// <param name="shape"></param>
        /// <returns></returns>
        public static IList<string> GetNames(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            int user_prop_row_count = UserDefinedCellHelper.GetCount(shape);

            if (user_prop_row_count < 1)
            {
                return new List<string>(0);
            }

            var prop_names = new List<string>(user_prop_row_count);
            var prop_section = shape.Section[UserDefinedCellHelper._userdefinedcell_section];
            var query_names = prop_section.ToEnumerable().Select(row => row.NameU);
            prop_names.AddRange(query_names);

            if (user_prop_row_count != prop_names.Count)
            {
                throw new AutomationException("Unexpected number of user-defined-cell names");
            }

            return prop_names;
        }

        public static bool IsValidName(string name)
        {
            if (name == null)
            {
                return false;
            }

            if (name.Length < 1)
            {
                return false;
            }

            const string space = " ";
            const string tab = "\t";
            const string carriage_return = "\r";
            const string line_feed = "\n";
            if (name.Contains(space) || name.Contains(tab) || name.Contains(carriage_return) || name.Contains(line_feed))
            {
                return false;
            }

            return true;
        }

        internal static void CheckValidName(string name)
        {
            if (!UserDefinedCellHelper.IsValidName(name))
            {
                string msg = string.Format("Invalid Name for User-Defined Cell: \"{0}\"", name);
                throw new AutomationException(msg);
            }
        }

        public static bool Contains(IVisio.Shape shape, string name)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            UserDefinedCellHelper.CheckValidName(name);

            string full_prop_name = UserDefinedCellHelper.GetRowName(name);

            var exists = (short)IVisio.VisExistsFlags.visExistsAnywhere;
            return 0 != (shape.CellExistsU[full_prop_name, exists]);
        }
    }
}