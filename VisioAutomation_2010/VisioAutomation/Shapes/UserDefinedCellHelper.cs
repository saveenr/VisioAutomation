using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public static class UserDefinedCellHelper
    {
        private static readonly short _udcell_section = ShapeSheet.SrcConstants.UserDefCellPrompt.Section;

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
            shape.DeleteRow(_udcell_section, row);
        }

        public static void Set(IVisio.Shape shape, string name, string value, string prompt)
        {
            var cells = new UserDefinedCellCells();
            cells.Value = value;
            cells.Prompt = prompt;
            cells.EncodeValues();
            Set(shape, name, cells);
        }

        public static void Set(IVisio.Shape shape, string name, UserDefinedCellCells cells)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            if (cells == null)
            {
                throw new System.ArgumentNullException(nameof(cells));
            }

            UserDefinedCellHelper.CheckValidName(name);

            if (UserDefinedCellHelper.Contains(shape, name))
            {
                // The user-defined cell already exists
                string full_prop_name = UserDefinedCellHelper.GetRowName(name);

                if (cells.Value.HasValue)
                {
                    string value_cell_name = full_prop_name;
                    var cell = shape.CellsU[value_cell_name];
                    cell.FormulaU = cells.Value.Value;
                }

                if (cells.Prompt.HasValue)
                {
                    string prompt_cell_name = full_prop_name + ".Prompt";
                    var cell = shape.CellsU[prompt_cell_name];
                    cell.FormulaU = cells.Prompt.Value;
                }
            }
            else
            {
                // The user-defined cell doesn't already exist
                short row = shape.AddNamedRow(_udcell_section, name, (short)IVisio.VisRowIndices.visRowUser);
                var src_value = new ShapeSheet.Src(_udcell_section, row, (short)IVisio.VisCellIndices.visUserValue);
                var src_prompt = new ShapeSheet.Src(_udcell_section, row, (short)IVisio.VisCellIndices.visUserPrompt);

                var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();

                if (cells.Value.HasValue)
                {
                    writer.SetFormula(src_value, cells.Value.Value);
                }

                if (cells.Prompt.HasValue)
                {
                    writer.SetFormula(src_prompt, cells.Prompt.Value);
                }

                writer.Commit(shape);            
            }
        }

        public static Dictionary<string, UserDefinedCellCells> GetDictionary(IVisio.Shape shape, ShapeSheet.CellValueType type)
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

            var  shape_data = UserDefinedCellHelper.GetUserDefinedCellCells(shape, type);

            var dic = new Dictionary<string,UserDefinedCellCells>(prop_count);
            for (int i = 0; i < prop_count; i++)
            {
                dic[prop_names[i]] = shape_data[i];
            }
            return dic;
        }

        public static List<Dictionary<string, UserDefinedCellCells>> GetDictionary(IVisio.Page page, IList<IVisio.Shape> shapes, ShapeSheet.CellValueType type)
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

            var list_list_customprops = UserDefinedCellHelper.GetUserDefinedCellCells(page,shapeids, CellValueType.Formula);

            var list_dic_customprops = new List<Dictionary<string, UserDefinedCellCells>>(shapeids.Count);

            for (int shape_index = 0; shape_index < shapes.Count; shape_index++)
            {
                var shape = shapes[shape_index];
                var list_customprops = list_list_customprops[shape_index];
                var prop_names = UserDefinedCellHelper.GetNames(shape);

                var dic_customprops = new Dictionary<string, UserDefinedCellCells>(list_customprops.Count);
                list_dic_customprops.Add(dic_customprops);
                for (int i = 0; i < list_customprops.Count ; i++)
                {
                    var prop_name = prop_names[i];
                    dic_customprops[prop_name] = list_customprops[i];
                }
            }

            return list_dic_customprops;
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
            if (0 == shape.SectionExists[_udcell_section, (short)IVisio.VisExistsFlags.visExistsAnywhere])
            {
                return 0;
            }

            var section = shape.Section[_udcell_section];

            if (section == null)
            {
                string msg = string.Format("Could not find the user-defined section for shape {0}", shape.NameU);
                throw new InternalAssertionException(msg);
            }

            int row_count = section.Shape.RowCount[_udcell_section];

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
            var prop_section = shape.Section[UserDefinedCellHelper._udcell_section];
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

        public static List<List<UserDefinedCellCells>> GetUserDefinedCellCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = UserDefinedCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<UserDefinedCellCells> GetUserDefinedCellCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = UserDefinedCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<UserDefinedCellCellsReader> UserDefinedCells_lazy_reader = new System.Lazy<UserDefinedCellCellsReader>();




        class UserDefinedCellCellsReader : CellGroupReader<UserDefinedCellCells>
        {
            public SectionQueryColumn Value { get; set; }
            public SectionQueryColumn Prompt { get; set; }

            public UserDefinedCellCellsReader() : base(new VisioAutomation.ShapeSheet.Query.SectionsQuery())
            {
                var sec = this.query_multirow.SectionQueries.Add(IVisio.VisSectionIndices.visSectionUser);
                this.Value = sec.Columns.Add(SrcConstants.UserDefCellValue, nameof(this.Value));
                this.Prompt = sec.Columns.Add(SrcConstants.UserDefCellPrompt, nameof(this.Prompt));
            }

            public override UserDefinedCellCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new UserDefinedCellCells();
                cells.Value = row[this.Value];
                cells.Prompt = row[this.Prompt];
                return cells;
            }
        }

    }
}