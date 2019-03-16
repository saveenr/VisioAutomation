using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

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

            string full_udcell_name = UserDefinedCellHelper.GetRowName(name);

            short row = shape.CellsU[full_udcell_name].Row;
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
                string full_udcell_name = UserDefinedCellHelper.GetRowName(name);

                if (cells.Value.HasValue)
                {
                    string value_cell_name = full_udcell_name;
                    var cell = shape.CellsU[value_cell_name];
                    cell.FormulaU = cells.Value.Value;
                }

                if (cells.Prompt.HasValue)
                {
                    string prompt_cell_name = full_udcell_name + ".Prompt";
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

        public static UserDefinedCellDictionary GetDictionary(IVisio.Shape shape, VASS.CellValueType type)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var udcell_count = UserDefinedCellHelper.GetCount(shape);
            if (udcell_count < 1)
            {
                return new UserDefinedCellDictionary(0);
            }

            var udcell_names = UserDefinedCellHelper.GetNames(shape);
            if (udcell_names.Count != udcell_count)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Unexpected number of user-define cell names");
            }

            var  shape_data = UserDefinedCellHelper.GetUserDefinedCellCells(shape, type);

            var dic = new UserDefinedCellDictionary(udcell_count);
            for (int i = 0; i < udcell_count; i++)
            {
                dic[udcell_names[i]] = shape_data[i];
            }
            return dic;
        }

        public static List<UserDefinedCellDictionary> GetDictionary(IVisio.Page page, IList<IVisio.Shape> shapes, VASS.CellValueType type)
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

            var list_list_udcells = UserDefinedCellHelper.GetUserDefinedCellCells(page,shapeids, VASS.CellValueType.Formula);

            var list_dics = new List<UserDefinedCellDictionary>(shapeids.Count);

            for (int shape_index = 0; shape_index < shapes.Count; shape_index++)
            {
                var shape = shapes[shape_index];
                var list_udcells = list_list_udcells[shape_index];
                var udcell_names = UserDefinedCellHelper.GetNames(shape);

                var dic_udcells = new UserDefinedCellDictionary(list_udcells.Count);
                list_dics.Add(dic_udcells);
                for (int i = 0; i < list_udcells.Count ; i++)
                {
                    var udcell_name = udcell_names[i];
                    dic_udcells[udcell_name] = list_udcells[i];
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

            // If the User-defined cell section does not exist then return zero immediately
            if (0 == shape.SectionExists[_udcell_section, (short)IVisio.VisExistsFlags.visExistsAnywhere])
            {
                return 0;
            }

            var section = shape.Section[_udcell_section];

            if (section == null)
            {
                string msg = string.Format("Could not find the user-defined section for shape {0}", shape.NameU);
                throw new VisioAutomation.Exceptions.InternalAssertionException(msg);
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

            int udcell_count = UserDefinedCellHelper.GetCount(shape);

            if (udcell_count < 1)
            {
                return new List<string>(0);
            }

            var udcell_names = new List<string>(udcell_count);
            var udcell_section = shape.Section[UserDefinedCellHelper._udcell_section];
            var query_names = udcell_section.ToEnumerable().Select(row => row.NameU);
            udcell_names.AddRange(query_names);

            if (udcell_count != udcell_names.Count)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Unexpected number of user-defined-cell names");
            }

            return udcell_names;
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

            string full_udcell_name = UserDefinedCellHelper.GetRowName(name);

            var exists = (short)IVisio.VisExistsFlags.visExistsAnywhere;
            return 0 != (shape.CellExistsU[full_udcell_name, exists]);
        }

        public static List<List<UserDefinedCellCells>> GetUserDefinedCellCells(IVisio.Page page, IList<int> shapeids, VASS.CellValueType type)
        {
            var reader = UserDefinedCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<UserDefinedCellCells> GetUserDefinedCellCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = UserDefinedCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<UserDefinedCellCellsBuilder> UserDefinedCells_lazy_builder = new System.Lazy<UserDefinedCellCellsBuilder>();




        class UserDefinedCellCellsBuilder : VASS.CellGroups.CellGroupBuilder<UserDefinedCellCells>
        {

            public UserDefinedCellCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.MultiRow)
            {
            }

            
            public override UserDefinedCellCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new UserDefinedCellCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }


                cells.Value = getcellvalue(nameof(UserDefinedCellCells.Value));
                cells.Prompt = getcellvalue(nameof(UserDefinedCellCells.Prompt));



                return cells;
            }
        }

    }
}