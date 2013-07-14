using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.UserDefinedCells
{
    public static class UserDefinedCellsHelper
    {
        private static short _userdefinedcell_section = VA.ShapeSheet.SRCConstants.User_Prompt.Section;

        private static string GetRowName(string name)
        {
            return "User." + name;
        }

        public static void Delete(IVisio.Shape shape, string name)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            CheckValidName(name);

            string full_prop_name = GetRowName(name);

            short row = shape.CellsU[full_prop_name].Row;
            shape.DeleteRow(_userdefinedcell_section, row);
        }

        public static void Update(IVisio.Shape shape, string name, string val)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            CheckValidName(name);

            if (val == null)
            {
                throw new System.ArgumentNullException("val");
            }

            if (!Contains(shape, name))
            {
                throw new AutomationException("user Property does not exist");
            }

            string full_prop_name = GetRowName(name);

            var cell = shape.CellsU[full_prop_name];

            if (cell == null)
            {
                string msg = string.Format("Could not retrieve cell for user property \"{0}\"", full_prop_name);
                throw new AutomationException(msg);
            }

            var update = new VA.ShapeSheet.Update();
            var src = new VA.ShapeSheet.SRC(_userdefinedcell_section, cell.Row, (short)IVisio.VisCellIndices.visUserValue);
            update.SetFormula(src, val);

            update.Execute(shape);
        }

        public static void Set(IVisio.Shape shape, string name, string value, string prompt)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            CheckValidName(name);

            if (Contains(shape, name))
            {
                Delete(shape, name);
            }

            short row = shape.AddNamedRow(
                _userdefinedcell_section,
                name,
                (short)IVisio.VisRowIndices.visRowUser);

            var update = new VA.ShapeSheet.Update();

            if (value != null)
            {
                string value_formula = Convert.StringToFormulaString(value);
                var src = new VA.ShapeSheet.SRC(_userdefinedcell_section, row, (short)IVisio.VisCellIndices.visUserValue);
                update.SetFormula(src, value_formula);
            }

            if (prompt != null)
            {
                string prompt_formula = Convert.StringToFormulaString(prompt);
                var src = new VA.ShapeSheet.SRC(_userdefinedcell_section, row, (short)IVisio.VisCellIndices.visUserPrompt);
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
        public static IList<UserDefinedCell> Get(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var prop_count = GetCount(shape);
            if (prop_count < 1)
            {
                return new List<UserDefinedCell>(0);
            }

            var prop_names = GetNames(shape);
            if (prop_names.Count != prop_count)
            {
                throw new AutomationException("Unexpected number of prop names");
            }

            var data = UserDefinedCell.queryex.GetFormulas(shape);

            if (data.SectionCells != null)
            {
                var sec = data.SectionCells[0];
                var custom_props = create_userdefined_cell_list(prop_names, sec);
                return custom_props;
            }
            else
            {
                return new List<UserDefinedCell>(0);
            }

        }

        public static IList<List<UserDefinedCell>> Get(IVisio.Page page, IList<IVisio.Shape> shapes)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            if (shapes == null)
            {
                throw new System.ArgumentNullException("shapes");
            }

            var shapeids = shapes.Select(s => s.ID).ToList();

            var data = UserDefinedCell.queryex.GetFormulas(page, shapeids);

            var custom_props = new List<List<UserDefinedCell>>(shapeids.Count);

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var shape_data = data[i];
                if (shape_data.SectionCells.Count > 0)
                {
                    var section_data = shape_data.SectionCells[0];
                    var prop_names = GetNames(shape);
                    var ud_cells = create_userdefined_cell_list(prop_names, section_data);
                    custom_props.Add(ud_cells);
                }
                else
                {
                    custom_props.Add(new List<UserDefinedCell>(0));
                    
                }
            }

            return custom_props;
        }

        public static List<UserDefinedCell> create_userdefined_cell_list(
            IList<string> prop_names,
            VA.ShapeSheet.Query.QueryEx.SectionResult<string> sectiondata)
        {
            var custom_props = new List<UserDefinedCell>();
            int name_index = 0;

            foreach (var prop_name in prop_names)
            {
                var custom_prop = new UserDefinedCell(prop_name);
                custom_prop.Value = sectiondata.Rows[name_index][UserDefinedCell.queryex.Value];
                custom_prop.Prompt = sectiondata.Rows[name_index][UserDefinedCell.queryex.Prompt];
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
        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            // If the User Property section does not exist then return zero immediately
            if (0 == shape.SectionExists[_userdefinedcell_section, (short)IVisio.VisExistsFlags.visExistsAnywhere])
            {
                return 0;
            }

            var section = shape.Section[_userdefinedcell_section];

            if (section == null)
            {
                string msg = string.Format("Could not find the user-defined section for shape {0}", shape.NameU);
                throw new AutomationException(msg);
            }

            int row_count = section.Shape.RowCount[_userdefinedcell_section];

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
                throw new System.ArgumentNullException("shape");
            }

            int user_prop_row_count = GetCount(shape);

            if (user_prop_row_count < 1)
            {
                return new List<string>(0);
            }

            var prop_names = new List<string>(user_prop_row_count);
            var prop_section = shape.Section[_userdefinedcell_section];
            var query_names = prop_section.AsEnumerable().Select(row => row.NameU);
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

            if (name.Contains(" ") || name.Contains("\t") || name.Contains("\r") || name.Contains("\n"))
            {
                return false;
            }

            return true;
        }

        internal static void CheckValidName(string name)
        {
            if (!IsValidName(name))
            {
                string msg = string.Format("Invalid Name for User-Defined Cell: \"{0}\"", name);
                throw new VA.AutomationException(msg);
            }
        }

        public static bool Contains(IVisio.Shape shape, string name)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            CheckValidName(name);

            string full_prop_name = GetRowName(name);

            var exists = (short)IVisio.VisExistsFlags.visExistsAnywhere;
            return 0 != (shape.CellExistsU[full_prop_name, exists]);
        }
    }
}