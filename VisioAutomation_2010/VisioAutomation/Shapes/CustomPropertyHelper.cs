using System;
using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public static class CustomPropertyHelper
    {
        static short vis_sec_prop = (short)IVisio.VisSectionIndices.visSectionProp;

        public static void Set(
            IVisio.Shape shape,
            string name,
            CustomPropertyCells cp)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            CustomPropertyHelper.CheckValidCustomPropertyName(name);

            if (CustomPropertyHelper.Contains(shape, name))
            {
                string full_prop_name = CustomPropertyHelper.GetRowName(name);
                var cell_propname = shape.CellsU[full_prop_name];

                if (cell_propname == null)
                {
                    string msg = string.Format("Could not retrieve cell for custom property \"{0}\"", full_prop_name);
                    throw new Exceptions.InternalAssertionException(msg);
                }

                var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
                writer.SetValues(cp, cell_propname.Row);

                writer.CommitFormulas(shape);

                return;
            }

            short row = shape.AddNamedRow(
                vis_sec_prop,
                name,
                (short)IVisio.VisRowIndices.visRowProp);

            CustomPropertyHelper.Set(shape, row, cp);
        }

        public static void Set(IVisio.Shape shape, short row, CustomPropertyCells cp)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(cp, row);

            writer.CommitFormulas(shape);
        }

        public static CustomPropertyDictionary GetDictionary(IVisio.Shape shape, VASS.CellValueType type)
        {
            var pairs = GetPairs(shape, type);
            var shape_custprop_dic = CustomPropertyDictionary.FromPairs(pairs);
            return shape_custprop_dic;
        }

        public static List<CustomPropertyDictionary> GetDictionary(IVisio.Page page, IList<IVisio.Shape> shapes, VASS.CellValueType type)
        {
            var shapeidpairs = ShapeSheet.Query.ShapeIdPairs.Create(shapes);
            var listof_listof_custpropscells = CustomPropertyCells.GetCells(page, shapeidpairs, type);
            var listof_custpropdics = GetListOfCpDic(shapeidpairs, listof_listof_custpropscells);

            return listof_custpropdics;
        }

        public static List<CustomPropertyCells> GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var shape_custprop_cells = CustomPropertyCells.GetCells(shape, type);
            return shape_custprop_cells;
        }

        public static List<List<CustomPropertyCells>> GetCells(IVisio.Page page, IList<IVisio.Shape> shapes, VASS.CellValueType type)
        {
            var shapeidpairs = ShapeSheet.Query.ShapeIdPairs.Create(shapes);
            var listof_listof_custpropscells = CustomPropertyCells.GetCells(page, shapeidpairs, type);
            return listof_listof_custpropscells;
        }

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            var exists_flag = (short)IVisio.VisExistsFlags.visExistsAnywhere;

            // If the Custom Property section does not exist then return zero immediately
            if (0 == shape.SectionExists[vis_sec_prop, exists_flag])
            {
                return 0;
            }

            var section = shape.Section[vis_sec_prop];

            if (section == null)
            {
                throw new System.NullReferenceException(nameof(section));
            }

            int row_count = section.Shape.RowCount[vis_sec_prop];

            return row_count;
        }

        public static List<string> GetNames(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            int custom_prop_row_count = CustomPropertyHelper.GetCount(shape);

            if (custom_prop_row_count < 1)
            {
                return new List<string>(0);
            }

            var prop_names = new List<string>(custom_prop_row_count);
            var prop_section = shape.Section[vis_sec_prop];
            var query_names = prop_section.ToEnumerable().Select(row => row.NameU);
            prop_names.AddRange(query_names);

            if (custom_prop_row_count != prop_names.Count)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Unexpected number of property names");
            }

            return prop_names;
        }

        private static bool IsValidName(string name, out string errmsg)
        {
            if (name == null)
            {
                errmsg = "The Custom Property name cannot be null";
                return false;
            }

            if (name.Length < 1)
            {
                errmsg = "The Custom Property name cannot have zero length";
                return false;
            }

            if (name.Contains(" ") || name.Contains("\t") || name.Contains("\r") || name.Contains("\n"))
            {
                errmsg = "The Custom Property name cannot contain any whitespace";
                return false;
            }

            if (name.StartsWith("Prop."))
            {
                errmsg = "The Custom Property name cannot begin with \"Prop.\".";
                return false;
            }

            errmsg = null;
            return true;
        }

        public static bool IsValidName(string name)
        {
            string errmsg;
            return CustomPropertyHelper.IsValidName(name, out errmsg);
        }

        internal static void CheckValidCustomPropertyName(string name)
        {
            string errmsg;
            if (!CustomPropertyHelper.IsValidName(name, out errmsg))
            {
                string msg = string.Format("Invalid Property Name: \"{0}\". {1}", name, errmsg);
                throw new System.ArgumentException(msg);
            }
        }

        public static bool Contains(IVisio.Shape shape, string name)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            CustomPropertyHelper.CheckValidCustomPropertyName(name);

            string full_prop_name = CustomPropertyHelper.GetRowName(name);

            var exists = (short)IVisio.VisExistsFlags.visExistsAnywhere;
            return 0 != (shape.CellExistsU[full_prop_name, exists]);
        }

        private static string GetRowName(string name)
        {
            return string.Format("Prop.{0}", name);
        }

        public static void Delete(IVisio.Shape shape, string name)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            CustomPropertyHelper.CheckValidCustomPropertyName(name);

            string full_prop_name = CustomPropertyHelper.GetRowName(name);

            short row = shape.CellsU[full_prop_name].Row;
            shape.DeleteRow(vis_sec_prop, row);
        }

        public static void Set(IVisio.Shape shape, string name, string value, int type)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            CustomPropertyHelper.CheckValidCustomPropertyName(name);

            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            // create a new property
            var cp = new CustomPropertyCells();
            cp.Value = value;
            cp.Type = type;

            CustomPropertyHelper.Set(shape, name, cp);
        }


        // ----------------------------------------
        // ----------------------------------------
        // ----------------------------------------
        // ----------------------------------------

        internal class CustomPropNameCellsPair
        {
            public string Name;
            public CustomPropertyCells Cells;

            public CustomPropNameCellsPair(string name, CustomPropertyCells cells)
            {
                this.Name = name;
                this.Cells = cells;
            }
        }

        private static List<CustomPropNameCellsPair> GetPairs(IVisio.Shape shape, VASS.CellValueType type)
        {
            var shape_custprop_cells = CustomPropertyCells.GetCells(shape, type);
            var shape_custprop_names = CustomPropertyHelper.GetNames(shape);
            var list = CreateListofPairs(shape_custprop_names, shape_custprop_cells);
            return list;
        }

        private static List<CustomPropNameCellsPair> CreateListofPairs(
            List<string> shape_custprop_names,
            List<CustomPropertyCells> shape_custprop_cells)
        {
            int num_props = shape_custprop_names.Count;

            var list = new List<CustomPropNameCellsPair>(num_props);
            var shape_custprop_indices = System.Linq.Enumerable.Range(0, num_props);
            foreach (int i in shape_custprop_indices)
            {
                string prop_name = shape_custprop_names[i];
                var shape_custprop_cell = shape_custprop_cells[i];
                var pair = new CustomPropNameCellsPair(prop_name, shape_custprop_cell);
                list.Add(pair);
            }

            return list;
        }

        private static List<CustomPropertyDictionary> GetListOfCpDic(
            ShapeSheet.Query.ShapeIdPairs shapeidpairs,
            List<List<CustomPropertyCells>> customprops_per_shape)
        {
            if (customprops_per_shape.Count != shapeidpairs.Count)
            {
                throw new Exceptions.InternalAssertionException();
            }

            var listof_listof_cppair = GetListOfCpPairLists(shapeidpairs, customprops_per_shape);
            var enumof_cpdic = listof_listof_cppair.Select(i => CustomPropertyDictionary.FromPairs(i));
            var list_cpdic = new List<CustomPropertyDictionary>(shapeidpairs.Count);
            list_cpdic.AddRange(enumof_cpdic);
            return list_cpdic;
        }

        private static List<List<CustomPropNameCellsPair>> GetListOfCpPairLists(
            ShapeSheet.Query.ShapeIdPairs shapeidpairs,
            List<List<CustomPropertyCells>> listof_listof_cpcells)
        {
            if (listof_listof_cpcells.Count != shapeidpairs.Count)
            {
                throw new Exceptions.InternalAssertionException();
            }

            var listof_listof_cppairs = new List<List<CustomPropNameCellsPair>>(shapeidpairs.Count);
            var shape_indices = System.Linq.Enumerable.Range(0, shapeidpairs.Count);

            foreach (int i in shape_indices)
            {
                var shape = shapeidpairs[i].Shape;
                var listof_cpnames = CustomPropertyHelper.GetNames(shape);
                var listof_cpcells = listof_listof_cpcells[i];

                int num_cps = listof_cpnames.Count;
                var cp_indices = Enumerable.Range(0, listof_cpnames.Count);
                var enumof_cppairs = cp_indices.Select(j => new CustomPropNameCellsPair(listof_cpnames[j], listof_cpcells[j])).ToList();
                listof_listof_cppairs.Add(enumof_cppairs);
            }
            return listof_listof_cppairs;
        }

    }
}