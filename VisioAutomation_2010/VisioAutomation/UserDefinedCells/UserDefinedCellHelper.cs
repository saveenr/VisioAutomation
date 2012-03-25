using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.UserDefinedCells
{
    public static class UserDefinedCellsHelper
    {
        public static string GetRowName(string name)
        {
            return "User." + name;
        }

        public static void DeleteUserDefinedCell(IVisio.Shape shape, string name)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            if (name == null)
            {
                throw new ArgumentNullException("name");
            }

            CheckValidUserDefinedCellName(name);

            string full_prop_name = GetRowName(name);

            short row = shape.CellsU[full_prop_name].Row;
            shape.DeleteRow(UserDefinedCell.query.Section, row);
        }

        public static void UpdateUserDefinedCell(IVisio.Shape shape, string name, string val)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            CheckValidUserDefinedCellName(name);

            if (val == null)
            {
                throw new ArgumentNullException("val");
            }

            if (!HasUserDefinedCell(shape, name))
            {
                throw new AutomationException("user Property does not exist");
            }

            string full_prop_name = GetRowName(name);

            var cell = shape.CellsU[full_prop_name];

            if (cell == null)
            {
                string msg = String.Format("Could not retrieve cell for user property \"{0}\"", full_prop_name);
                throw new AutomationException(msg);
            }

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            var src = new VA.ShapeSheet.SRC(UserDefinedCell.query.Section, cell.Row, (short)IVisio.VisCellIndices.visUserValue);
            update.SetFormula(src, val);

            update.Execute(shape);
        }

        public static void SetUserDefinedCell(IVisio.Shape shape, string name, string value, string prompt)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            CheckValidUserDefinedCellName(name);

            if (HasUserDefinedCell(shape, name))
            {
                DeleteUserDefinedCell(shape, name);
            }

            short row = shape.AddNamedRow(
                UserDefinedCell.query.Section,
                name,
                (short)IVisio.VisRowIndices.visRowUser);

            var update = new VA.ShapeSheet.Update.SRCUpdate();

            if (value != null)
            {
                string value_formula = Convert.StringToFormulaString(value);
                var src = new VA.ShapeSheet.SRC(UserDefinedCell.query.Section, row, (short)IVisio.VisCellIndices.visUserValue);
                update.SetFormula(src, value_formula);
            }

            if (prompt != null)
            {
                string prompt_formula = Convert.StringToFormulaString(prompt);
                var src = new VA.ShapeSheet.SRC(UserDefinedCell.query.Section, row, (short)IVisio.VisCellIndices.visUserPrompt);
                update.SetFormula(src, prompt_formula);
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
        public static IList<UserDefinedCell> GetUserDefinedCells(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            var prop_count = GetUserDefinedCellCount(shape);
            if (prop_count < 1)
            {
                return new List<UserDefinedCell>(0);
            }

            var prop_names = GetUserDefinedCellNames(shape);
            if (prop_names.Count != prop_count)
            {
                throw new AutomationException("Unexpected number of prop names");
            }

            var formulas = UserDefinedCell.query.GetFormulas(shape);

            var rows = new List<int>(formulas.RowCount);
            for (int row = 0; row < formulas.RowCount; row++)
            {
                rows.Add(row);
            }
            var custom_props = create_userdefined_cell_list(prop_names, formulas, rows);

            return custom_props;
        }

        public static IList<List<UserDefinedCell>> GetUserDefinedCells(IVisio.Page page, IList<IVisio.Shape> shapes)
        {
            if (page == null)
            {
                throw new ArgumentNullException("page");
            }

            if (shapes == null)
            {
                throw new ArgumentNullException("shapes");
            }

            var shapeids = shapes.Select(s => s.ID).ToList();

            var formulas = UserDefinedCell.query.GetFormulas(page, shapeids);

            var custom_props = new List<List<UserDefinedCell>>(shapeids.Count);

            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var group = formulas.Groups[shape_index];
                var shape = shapes[shape_index];
                var prop_names = GetUserDefinedCellNames(shape);
                var custom_props_for_shape = create_userdefined_cell_list(prop_names, formulas, group.RowIndices.ToList());
                custom_props.Add(custom_props_for_shape);
            }

            return custom_props;
        }

        public static List<UserDefinedCell> create_userdefined_cell_list(
            IList<string> prop_names,
            VA.ShapeSheet.Data.Table<string> formulas,
            IList<int> rows)
        {
            var custom_props = new List<UserDefinedCell>();
            int name_index = 0;

            foreach (int row in rows)
            {
                var prop_name = prop_names[name_index];

                var custom_prop = new UserDefinedCell(prop_name);
                custom_prop.Value = formulas[row, UserDefinedCell.query.Value];
                custom_prop.Prompt = formulas[row, UserDefinedCell.query.Prompt];
                custom_props.Add(custom_prop);

                name_index++;
            }

            return custom_props;
        }

        /// <summary>
        /// Get the number of user-defined cells for the shape.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        public static int GetUserDefinedCellCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            // If the User Property section does not exist then return zero immediately
            if (0 == shape.SectionExists[UserDefinedCell.query.Section, (short)IVisio.VisExistsFlags.visExistsAnywhere])
            {
                return 0;
            }

            var section = shape.Section[UserDefinedCell.query.Section];

            if (section == null)
            {
                string msg = String.Format("Could not find the user-defined section for shape {0}", shape.NameU);
                throw new AutomationException(msg);
            }

            int row_count = section.Shape.RowCount[UserDefinedCell.query.Section];

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
        public static IList<string> GetUserDefinedCellNames(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            int user_prop_row_count = GetUserDefinedCellCount(shape);

            if (user_prop_row_count < 1)
            {
                return new List<string>(0);
            }

            var prop_names = new List<string>(user_prop_row_count);
            var prop_section = shape.Section[UserDefinedCell.query.Section];
            var query_names = prop_section.AsEnumerable().Select(row => row.NameU);
            prop_names.AddRange(query_names);

            if (user_prop_row_count != prop_names.Count)
            {
                throw new AutomationException("Unexpected number of property names");
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

            return true;
        }

        public static void CheckValidName(string name)
        {
            if (!IsValidName(name))
            {
                throw new System.ArgumentException("name");                
            }
        }

        public static bool IsValidUserDefinedCellName(string name)
        {
            if (name == null)
            {
                return false;
            }

            if (name.Length < 1)
            {
                return false;
            }

            if (name.Contains(" ") || name.Contains("\t") || name.Contains("\r") || name.Contains("\n"))
            {
                return false;
            }

            return true;
        }

        public static void CheckValidUserDefinedCellName(string name)
        {
            if (!IsValidUserDefinedCellName(name))
            {
                string msg = String.Format("Invalid Property Name: \"{0}\"", name);
                throw new VA.AutomationException(msg);
            }
        }

        public static bool HasUserDefinedCell(IVisio.Shape shape, string name)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            CheckValidUserDefinedCellName(name);

            string full_prop_name = GetRowName(name);

            var exists = (short)IVisio.VisExistsFlags.visExistsAnywhere;
            return 0 != (shape.CellExistsU[full_prop_name, exists]);
        }
    }
}