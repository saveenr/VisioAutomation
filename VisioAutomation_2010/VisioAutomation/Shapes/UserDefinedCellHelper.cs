using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public static class UserDefinedCellHelper
    {
        private static readonly short _userdefinedcell_section = ShapeSheet.SrcConstants.UserDefCellPrompt.Section;

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

        public static void Set(IVisio.Shape shape, string name, ShapeSheet.CellValueLiteral udfcell_value, ShapeSheet.CellValueLiteral udfcell_prompt)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            UserDefinedCellHelper.CheckValidName(name);

            if (UserDefinedCellHelper.Contains(shape, name))
            {
                string full_prop_name = UserDefinedCellHelper.GetRowName(name);

                if (udfcell_value.HasValue)
                {
                    string value_cell_name = full_prop_name;
                    var cell = shape.CellsU[value_cell_name];
                    cell.FormulaU = VisioAutomation.Utilities.Convert.StringToFormulaString(udfcell_value.Value);                    
                }

                if (udfcell_prompt.HasValue)
                {
                    string prompt_cell_name = full_prop_name+".Prompt";
                    var cell = shape.CellsU[prompt_cell_name];
                    cell.FormulaU = VisioAutomation.Utilities.Convert.StringToFormulaString(udfcell_prompt.Value);                                        
                }
                return;
            }

            short row = shape.AddNamedRow(
                UserDefinedCellHelper._userdefinedcell_section,
                name,
                (short)IVisio.VisRowIndices.visRowUser);

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();

            if (udfcell_value.HasValue)
            {
                var src = new ShapeSheet.Src(UserDefinedCellHelper._userdefinedcell_section, row, (short)IVisio.VisCellIndices.visUserValue);
                var formula = VisioAutomation.Utilities.Convert.StringToFormulaString(udfcell_value.Value);
                writer.SetFormula(src, formula);
            }

            if (udfcell_prompt.HasValue)
            {
                var src = new ShapeSheet.Src(UserDefinedCellHelper._userdefinedcell_section, row, (short)IVisio.VisCellIndices.visUserPrompt);
                var formula = VisioAutomation.Utilities.Convert.StringToFormulaString(udfcell_prompt.Value);
                writer.SetFormula(src, formula);
            }

            writer.Commit(shape);
        }


        public static Dictionary<string, UserDefinedCellCells> GetDictionary(IVisio.Shape shape, ShapeSheet.CellValueType cvt)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var prop_count = UserDefinedCellHelper.GetCount(shape);
            if (prop_count < 1)
            {
                return new Dictionary<string, UserDefinedCellCells>(0);
            }

            var prop_names = UserDefinedCellHelper.GetNames(shape);
            if (prop_names.Count != prop_count)
            {
                throw new InternalAssertionException("Unexpected number of prop names");
            }

            var  shape_data = UserDefinedCellCells.GetCells(shape, cvt);

            var dic = new Dictionary<string,UserDefinedCellCells>(prop_count);
            for (int i = 0; i < prop_count; i++)
            {
                dic[prop_names[i]] = shape_data[i];
            }
            return dic;
        }

        public static List<Dictionary<string, UserDefinedCellCells>> GetDictionary(IVisio.Page page, IList<IVisio.Shape> shapes, ShapeSheet.CellValueType cvt)
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

            var list_data = UserDefinedCellCells.GetCells(page,shapeids, CellValueType.Formula);

            var list_dics = new List<Dictionary<string, UserDefinedCellCells>>(shapeids.Count);

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var shape_data = list_data[i];
                var prop_names = UserDefinedCellHelper.GetNames(shape);

                var dic = new Dictionary<string, UserDefinedCellCells>(shape_data.Count);
                list_dics.Add(dic);
                for (int j = 0; j < shape_data.Count ; j++)
                {
                    dic[prop_names[j]] = shape_data[j];
                }
            }

            return list_dics;
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
                throw new InternalAssertionException(msg);
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
        public static List<string> GetNames(IVisio.Shape shape)
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
                throw new InternalAssertionException("Unexpected number of user-defined-cell names");
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

        public static void CheckValidName(string name)
        {
            if (!UserDefinedCellHelper.IsValidName(name))
            {
                string msg = string.Format("Invalid Name for User-Defined Cell: \"{0}\"", name);
                throw new System.ArgumentException(msg);
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