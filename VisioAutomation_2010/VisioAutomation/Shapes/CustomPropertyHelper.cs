using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public static class CustomPropertyHelper
    {
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
                    throw new InternalAssertionException(msg);
                }

                var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
                cp.SetFormulas(writer, cell_propname.Row);

                writer.Commit(shape);

                return;
            }

            short row = shape.AddNamedRow(
                (short)IVisio.VisSectionIndices.visSectionProp,
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
            cp.SetFormulas(writer, row);

            writer.Commit(shape);
        }

        /// <summary>
        /// Gets all the custom properties defined on a shape
        /// </summary>
        /// <remarks>
        /// If there are no custom properties then null will be returned</remarks>
        /// <param name="shape"></param>
        /// <returns>A list of custom properties</returns>
        public static CustomPropertyDictionary Get(IVisio.Shape shape)
        {
            var prop_names = CustomPropertyHelper.GetNames(shape);
            var dic = new CustomPropertyDictionary(prop_names.Count);
            var cells = CustomPropertyCells.GetCells(shape);

            for (int prop_index = 0; prop_index < prop_names.Count(); prop_index++)
            {
                string prop_name = prop_names[prop_index];
                dic[prop_name] = cells[prop_index];
            }

            return dic;
        }

        public static List<CustomPropertyDictionary> Get(IVisio.Page page, IList<IVisio.Shape> shapes)
        {
            if (page == null)
            {
                throw new ArgumentNullException(nameof(page));
            }

            if (shapes == null)
            {
                throw new ArgumentNullException(nameof(shapes));
            }

            var shapeids = shapes.Select(s => s.ID).ToList();
            var customprops_dic = new List<CustomPropertyDictionary>(shapeids.Count);
            var customprops_per_shape = CustomPropertyCells.GetCells(page, shapeids);
            
            if (customprops_per_shape.Count!=shapeids.Count)
            {
                throw new InternalAssertionException();
            }

            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shape = shapes[shape_index];
                var customprops_for_shape = customprops_per_shape[shape_index];
                var prop_names = CustomPropertyHelper.GetNames(shape);

                if (customprops_for_shape.Count != prop_names.Count)
                {
                    throw new InternalAssertionException();
                }

                var dic = new CustomPropertyDictionary(prop_names.Count);
                
                for (int prop_index=0; prop_index< prop_names.Count(); prop_index++)
                {
                    string prop_name = prop_names[prop_index];
                    dic[prop_name] = customprops_for_shape[prop_index];
                }

                customprops_dic.Add(dic);
            }

            return customprops_dic;
        }

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            // If the Custom Property section does not exist then return zero immediately
            if (0 == shape.SectionExists[(short)IVisio.VisSectionIndices.visSectionProp, (short)IVisio.VisExistsFlags.visExistsAnywhere])
            {
                return 0;
            }

            var section = shape.Section[(short)IVisio.VisSectionIndices.visSectionProp];

            if (section == null)
            {
                throw new System.NullReferenceException(nameof(section));
            }

            int row_count = section.Shape.RowCount[(short)IVisio.VisSectionIndices.visSectionProp];

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
            var prop_section = shape.Section[(short)IVisio.VisSectionIndices.visSectionProp];
            var query_names = prop_section.ToEnumerable().Select(row => row.NameU);
            prop_names.AddRange(query_names);

            if (custom_prop_row_count != prop_names.Count)
            {
                throw new InternalAssertionException("Unexpected number of property names");
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
            shape.DeleteRow((short)IVisio.VisSectionIndices.visSectionProp, row);
        }

        public static void Set(IVisio.Shape shape, string name, string val)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            CustomPropertyHelper.CheckValidCustomPropertyName(name);

            if (val == null)
            {
                throw new ArgumentNullException(nameof(val));
            }

            // create a new property
            var cp = new CustomPropertyCells();
            cp.Value = val;
            cp.Type = 0; // 0 = string

            CustomPropertyHelper.Set(shape, name, cp);
        }
    }
}