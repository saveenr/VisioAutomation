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

            string full_udcell_name = UserDefinedCellHelper.__GetRowName(name);

            short row = shape.CellsU[full_udcell_name].Row;
            shape.DeleteRow(_udcell_section, row);
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
                string full_udcell_name = UserDefinedCellHelper.__GetRowName(name);

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
                    writer.SetValue(src_value, cells.Value.Value);
                }

                if (cells.Prompt.HasValue)
                {
                    writer.SetValue(src_prompt, cells.Prompt.Value);
                }

                writer.Commit(shape, VASS.CellValueType.Formula);            
            }
        }

        public static UserDefinedCellDictionary GetDictionary(IVisio.Shape shape, VASS.CellValueType type)
        {
            var pairs = __GetPairs(shape, type);
            var dic = UserDefinedCellDictionary.FromPairs(pairs);
            return dic;
        }

        public static List<UserDefinedCellDictionary> GetDictionary(IVisio.Page page, Core.ShapeIDPairs shapeidpairs, VASS.CellValueType type)
        {
            int num_shapes = shapeidpairs.Count;
            var list_list_pair = __GetPairs(page, shapeidpairs, type);

            var list_dic = new List<UserDefinedCellDictionary>(num_shapes);
            var dics = list_list_pair.Select(list_pair => UserDefinedCellDictionary.FromPairs(list_pair));
            list_dic.AddRange(dics);

            return list_dic;
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

            string full_udcell_name = UserDefinedCellHelper.__GetRowName(name);

            var exists = (short)IVisio.VisExistsFlags.visExistsAnywhere;
            return 0 != (shape.CellExistsU[full_udcell_name, exists]);
        }

        // ---------------------------------------------------------------
        // ---------------------------------------------------------------
        // ---------------------------------------------------------------

        private static List<UserDefinedCellNameCellsPair> __CreateNamePairs(int shapeid, List<string> udcell_names, List<UserDefinedCellCells> list_udcells)
        {
            var namepairs = new List<UserDefinedCellNameCellsPair>(list_udcells.Count);
            int num_udcells = list_udcells.Count;
            var udcell_rows = Enumerable.Range(0, num_udcells);
            foreach (int udcell_row in udcell_rows)
            {
                var udcell_name = udcell_names[udcell_row];
                var pair = new UserDefinedCellNameCellsPair(shapeid, udcell_row, udcell_name, list_udcells[udcell_row]);

                namepairs.Add(pair);
            }

            return namepairs;
        }

        private static List<List<UserDefinedCellNameCellsPair>> __GetPairs(IVisio.Page page, Core.ShapeIDPairs shapeidpairs, VASS.CellValueType type)
        {
            var list_list_udcells = UserDefinedCellCells.GetCells(page, shapeidpairs, type);
            int num_shapes = shapeidpairs.Count;
            var list_list_pairs = new List<List<UserDefinedCellNameCellsPair>>(num_shapes);
            var shape_indices = Enumerable.Range(0, num_shapes);
            foreach (int shape_index in shape_indices)
            {
                var shapeidpair = shapeidpairs[shape_index];
                var udcell_names = UserDefinedCellHelper.GetNames(shapeidpair.Shape);
                var list_udcells = list_list_udcells[shape_index];
                var list_pairs = __CreateNamePairs(shapeidpair.ShapeID,udcell_names, list_udcells);


                list_list_pairs.Add(list_pairs);
            }

            return list_list_pairs;
        }

        private static List<UserDefinedCellNameCellsPair> __GetPairs(IVisio.Shape shape, VASS.CellValueType type)
        {
            var listof_udcellcells = UserDefinedCellCells.GetCells(shape, type);


            int num_udcells = listof_udcellcells.Count;

            var udcell_names = UserDefinedCellHelper.GetNames(shape);
            if (udcell_names.Count != num_udcells)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Unexpected number of user-define cell names");
            }

            int shapeid = shape.ID16;
            var pairs = new List<UserDefinedCellNameCellsPair>(num_udcells);
            var udcell_rows = Enumerable.Range(0, num_udcells);
            foreach (int udcell_row in udcell_rows)
            {
                var pair = new UserDefinedCellNameCellsPair(shapeid, udcell_row, udcell_names[udcell_row], listof_udcellcells[udcell_row]);
                pairs.Add(pair);
            }

            return pairs;
        }


        private static string __GetRowName(string name)
        {
            return "User." + name;
        }


    }
}