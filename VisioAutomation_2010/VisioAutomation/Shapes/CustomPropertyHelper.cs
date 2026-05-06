using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.CellRecords;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public static class CustomPropertyHelper
    {
        static short vis_sec_prop = (short) IVisio.VisSectionIndices.visSectionProp;

        public static void Set(
            IVisio.Shape shape,
            string name,
            CustomPropertyCells cp)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }
            if (cp == null)
            {
                throw new System.ArgumentNullException(nameof(cp));
            }

            __CheckValidCustomPropertyName(name);

            short row;
            if (Contains(shape, name))
            {
                string full_prop_name = __GetRowName(name);
                var cell_propname = shape.CellsU[full_prop_name];

                if (cell_propname == null)
                {
                    string msg = string.Format("Could not retrieve cell for custom property \"{0}\"", full_prop_name);
                    throw new Exceptions.InternalAssertionException(msg);
                }

                row = cell_propname.Row;
            }
            else
            {
                row = shape.AddNamedRow(
                    vis_sec_prop,
                    name,
                    (short) IVisio.VisRowIndices.visRowProp);
            }

            var writer = new ShapeSheet.Writers.SrcWriter();
            writer.SetValues(cp, row);
            __CommitWithFormulaErrorWrapping(writer, shape, name);
        }

        public static void Set(IVisio.Shape shape, short row, CustomPropertyCells cp)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var writer = new ShapeSheet.Writers.SrcWriter();
            writer.SetValues(cp, row);
            __CommitWithFormulaErrorWrapping(writer, shape, null);
        }

        // Issue #144 detect-and-rethrow. Visio raises COMException with message
        // "#NAME?" (or similar formula-error markers) when a cell formula is
        // invalid. The most common cause is an unencoded string assigned to a
        // formula-typed field. Wrap that case with a self-explanatory message
        // pointing at SetString/EncodeValues so the user knows what to do.
        private static void __CommitWithFormulaErrorWrapping(
            ShapeSheet.Writers.SrcWriter writer,
            IVisio.Shape shape,
            string propertyName)
        {
            try
            {
                writer.Commit(shape, Core.CellValueType.Formula);
            }
            catch (System.Runtime.InteropServices.COMException ex) when (__LooksLikeFormulaError(ex))
            {
                throw new System.ArgumentException(__BuildEncodingDiagnostic(propertyName, ex), ex);
            }
        }

        private static bool __LooksLikeFormulaError(System.Runtime.InteropServices.COMException ex)
        {
            string msg = ex.Message;
            if (string.IsNullOrEmpty(msg)) { return false; }
            return msg.Contains("#NAME?")
                || msg.Contains("#VALUE!")
                || msg.Contains("#REF!")
                || msg.Contains("#NUM!")
                || msg.Contains("#DIV/0!");
        }

        private static string __BuildEncodingDiagnostic(
            string propertyName,
            System.Runtime.InteropServices.COMException ex)
        {
            string ctx = string.IsNullOrEmpty(propertyName)
                ? "a custom property cell"
                : string.Format("custom property '{0}'", propertyName);
            return string.Format(
                "Visio rejected the formula for {0} ('{1}'). This typically means a CustomPropertyCells field (Formula, Label, Format, or Prompt) was assigned a raw string. Use SetString/SetNumber/SetBool/SetDate to set typed values, or call EncodeValues() before Set, to ensure values are valid Visio formulas.",
                ctx, ex.Message.Trim());
        }

        public static CustomPropertyDictionary GetDictionary(IVisio.Shape shape, Core.CellValueType type)


        {
            var pairs = __GetPairs(shape, type);
            var shape_custprop_dic = CustomPropertyDictionary.FromPairs(pairs);
            return shape_custprop_dic;
        }

        public static List<CustomPropertyDictionary> GetDictionary(IVisio.Page page, Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {
            var listof_listof_custpropscells = CustomPropertyCells.GetCells(page, shapeidpairs, type);
            var listof_custpropdics = _get_cells_as_list(shapeidpairs, listof_listof_custpropscells);

            return listof_custpropdics;
        }

        public static List<CustomPropertyDictionary> _get_cells_as_list(
            Core.ShapeIDPairs shapeidpairs,
            CellRecordsGroup<CustomPropertyCells> customprops_per_shape)
        {
            if (customprops_per_shape.Count != shapeidpairs.Count)
            {
                throw new Exceptions.InternalAssertionException();
            }

            var listof_listof_cppair = __GetListOfCpPairLists(shapeidpairs, customprops_per_shape);
            var enumof_cpdic = listof_listof_cppair.Select(i => CustomPropertyDictionary.FromPairs(i));
            var list_cpdic = new List<CustomPropertyDictionary>(shapeidpairs.Count);
            list_cpdic.AddRange(enumof_cpdic);
            return list_cpdic;
        }


        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var exists_flag = (short) IVisio.VisExistsFlags.visExistsAnywhere;

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
                throw new System.ArgumentNullException(nameof(shape));
            }

            int custom_prop_row_count = GetCount(shape);

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
                throw new Exceptions.InternalAssertionException("Unexpected number of property names");
            }

            return prop_names;
        }

        public static bool IsValidName(string name)
        {
            string errmsg;
            return __IsValidName(name, out errmsg);
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

            __CheckValidCustomPropertyName(name);

            string full_prop_name = __GetRowName(name);

            var exists = (short) IVisio.VisExistsFlags.visExistsAnywhere;
            return 0 != (shape.CellExistsU[full_prop_name, exists]);
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

            __CheckValidCustomPropertyName(name);

            string full_prop_name = __GetRowName(name);

            short row = shape.CellsU[full_prop_name].Row;
            shape.DeleteRow(vis_sec_prop, row);
        }

        public static void Set(IVisio.Shape shape, string name, string value, int type)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            __CheckValidCustomPropertyName(name);

            if (value == null)
            {
                throw new System.ArgumentNullException(nameof(value));
            }

            // create a new property
            var cp = new CustomPropertyCells();
            cp.Formula = value;
            cp.Type = type;

            Set(shape, name, cp);
        }


        // ----------------------------------------
        // ----------------------------------------
        // ----------------------------------------
        // ----------------------------------------

        private static List<CustomPropertyNameCellsPair> __GetPairs(IVisio.Shape shape, Core.CellValueType type)
        {
            var shape_custprop_cells = CustomPropertyCells.GetCells(shape, type);
            var shape_custprop_names = GetNames(shape);
            int shapeid = shape.ID16;
            var list = __CreateListofPairs(shape_custprop_names, shape_custprop_cells, shapeid);
            return list;
        }

        private static List<CustomPropertyNameCellsPair> __CreateListofPairs(
            List<string> shape_custprop_names,
            ShapeSheet.CellRecords.CellRecords<CustomPropertyCells> shape_custprop_cells,
            int shapeid)
        {
            int num_props = shape_custprop_names.Count;

            var list = new List<CustomPropertyNameCellsPair>(num_props);
            var custprop_rows = Enumerable.Range(0, num_props);
            foreach (int custprop_row in custprop_rows)
            {
                string prop_name = shape_custprop_names[custprop_row];
                var shape_custprop_cell = shape_custprop_cells[custprop_row];
                var pair = new CustomPropertyNameCellsPair(shapeid, custprop_row, prop_name, shape_custprop_cell);

                list.Add(pair);
            }

            return list;
        }

        private static List<List<CustomPropertyNameCellsPair>> __GetListOfCpPairLists(
            Core.ShapeIDPairs shapeidpairs,
            CellRecordsGroup<CustomPropertyCells> listof_listof_cpcells)
        {
            if (listof_listof_cpcells.Count != shapeidpairs.Count)
            {
                throw new Exceptions.InternalAssertionException();
            }

            var listof_listof_cppairs = new List<List<CustomPropertyNameCellsPair>>(shapeidpairs.Count);

            var shape_indices = Enumerable.Range(0, shapeidpairs.Count);

            foreach (int i in shape_indices)
            {
                var shape = shapeidpairs[i].Shape;
                var listof_cpnames = GetNames(shape);
                var listof_cpcells = listof_listof_cpcells[i];

                int num_cps = listof_cpnames.Count;

                var cp_indices = Enumerable.Range(0, num_cps);
                var listof_cppairs = new List<CustomPropertyNameCellsPair>(num_cps);
                foreach (int cprow in cp_indices)
                {
                    int shapeid = shapeidpairs[i].ShapeID;
                    var cppair =
                        new CustomPropertyNameCellsPair(shapeid, cprow, listof_cpnames[cprow], listof_cpcells[cprow]);
                    listof_cppairs.Add(cppair);
                }

                listof_listof_cppairs.Add(listof_cppairs);
            }

            return listof_listof_cppairs;
        }


        private static string __GetRowName(string name)
        {
            return string.Format("Prop.{0}", name);
        }

        private static bool __IsValidName(string name, out string errmsg)
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

        internal static void __CheckValidCustomPropertyName(string name)
        {
            string errmsg;
            if (!__IsValidName(name, out errmsg))
            {
                string msg = string.Format("Invalid Property Name: \"{0}\". {1}", name, errmsg);
                throw new System.ArgumentException(msg);
            }
        }
    }
}