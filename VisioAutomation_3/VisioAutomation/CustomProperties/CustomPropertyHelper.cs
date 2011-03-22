using System;
using Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.CustomProperties
{
    public static class CustomPropertyHelper
    {
        public static void UpdateCustomProperty(Shape shape, string name, string val)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            CheckValidCustomPropertyName(name);

            if (val == null)
            {
                throw new ArgumentNullException("val");
            }

            if (!HasCustomProperty(shape, name))
            {
                throw new AutomationException("Custom Property does not exist");
            }

            string full_prop_name = GetRowName(name);

            var cell_propname = shape.CellsU[full_prop_name];

            if (cell_propname == null)
            {
                string msg = String.Format("Could not retrieve cell for custom property \"{0}\"", full_prop_name);
                throw new AutomationException(msg);
            }

            short row = cell_propname.Row;
            var cell_propval =
                shape.CellsSRC[CustomPropertyCells.custprop_query.Section, row, (short)VisCellIndices.visCustPropsValue];
            cell_propval.FormulaU = val;
        }


        public static void SetCustomProperty(
            Shape shape,
            string name,
            VA.CustomProperties.CustomPropertyCells cp)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            CheckValidCustomPropertyName(name);

            if (HasCustomProperty(shape, name))
            {
                DeleteCustomProperty(shape, name);
            }

            short row = shape.AddNamedRow(
                CustomPropertyCells.custprop_query.Section,
                name,
                (short)VisRowIndices.visRowProp);

            SetCustomProperty(shape, row, cp);
        }

        public static void SetCustomProperty( Shape shape, short row, VA.CustomProperties.CustomPropertyCells cp)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            cp.Apply(update, row);
            update.Execute(shape);
        }

        /// <summary>
        /// Gets all the custom properties defined on a shape
        /// </summary>
        /// <remarks>
        /// If there are no custom properties then null will be returned</remarks>
        /// <param name="shape"></param>
        /// <returns>A list of custom properties</returns>
        public static IDictionary<string, CustomPropertyCells> GetCustomProperties(Shape shape)
        {
            var qds = CustomPropertyCells.custprop_query.GetFormulasAndResults<double>(shape);

            var prop_names = GetCustomPropertyNames(shape);
            if (prop_names.Count != qds.RowCount)
            {
                throw new AutomationException("Unexpected number of prop names");
            }

            var rows = new List<int>(qds.RowCount);
            for (int row = 0; row < qds.RowCount; row++)
            {
                rows.Add(row);
            }
            var custom_props = get_custom_props(prop_names, qds, rows);

            return custom_props;
        }

        public static IList<Dictionary<string, CustomPropertyCells>> GetCustomProperties(Page page, IList<Shape> shapes)
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

            var qds = CustomPropertyCells.custprop_query.GetFormulasAndResults<double>(page, shapeids);

            var customprops = new List<Dictionary<string, CustomPropertyCells>>(shapeids.Count);
            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var group = qds.Formulas.Groups[shape_index];
                var group_rows = group.RowIndices.ToList();
                var shape = shapes.ElementAt(shape_index);
                var prop_names = GetCustomPropertyNames(shape);
                var customprops_for_shape = get_custom_props(prop_names, qds, group_rows);
                customprops.Add(customprops_for_shape);
            }

            return customprops;
        }

        public static Dictionary<string, CustomPropertyCells> get_custom_props(IList<string> prop_names, VA.ShapeSheet.Query.QueryDataSet<double> qds, IList<int> group_rows)
        {
            if (prop_names.Count != group_rows.Count)
            {
                throw new AutomationException("Different number of names and result rows");
            }

            int num_props = prop_names.Count;
            var custom_properties = new Dictionary<string, CustomPropertyCells>(num_props);


            int prop_index = 0;
            foreach (int row in group_rows)
            {
                var prop_name = prop_names[prop_index];

                var cp = new CustomPropertyCells();

                cp.Value = qds.GetItem(row, CustomPropertyCells.custprop_query.Value);
                cp.Calendar = qds.GetItem(row, CustomPropertyCells.custprop_query.Calendar, v => (int)v);
                cp.Format = qds.GetItem(row, CustomPropertyCells.custprop_query.Format);
                cp.Invisible = qds.GetItem(row, CustomPropertyCells.custprop_query.Invis, v => (int)v);
                cp.Label = qds.GetItem(row, CustomPropertyCells.custprop_query.Label);
                cp.LangId = qds.GetItem(row, CustomPropertyCells.custprop_query.LangID, v => (int)v);
                cp.Prompt = qds.GetItem(row, CustomPropertyCells.custprop_query.Prompt);
                cp.SortKey = qds.GetItem(row, CustomPropertyCells.custprop_query.SortKey, v => (int)v);
                cp.Type = qds.GetItem(row, CustomPropertyCells.custprop_query.Type, v => (VA.CustomProperties.FormatShapeData)((int)v));

                custom_properties[prop_name] = cp;

                prop_index++;
            }

            return custom_properties;
        }

        public static int GetCustomPropertyCount(Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            // If the Custom Property section does not exist then return zero immediately
            if (0 == shape.SectionExists[CustomPropertyCells.custprop_query.Section, (short)VisExistsFlags.visExistsAnywhere])
            {
                return 0;
            }

            var section = shape.Section[CustomPropertyCells.custprop_query.Section];

            if (section == null)
            {
                throw new AutomationException("section is null");
            }

            int row_count = section.Shape.RowCount[CustomPropertyCells.custprop_query.Section];

            return row_count;
        }

        public static IList<string> GetCustomPropertyNames(Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            int custom_prop_row_count = GetCustomPropertyCount(shape);

            if (custom_prop_row_count < 1)
            {
                return new List<string>(0);
            }

            var prop_names = new List<string>(custom_prop_row_count);
            var prop_section = shape.Section[CustomPropertyCells.custprop_query.Section];
            var query_names = prop_section.AsEnumerable().Select(row => row.NameU);
            prop_names.AddRange(query_names);

            if (custom_prop_row_count != prop_names.Count)
            {
                throw new AutomationException("Unexpected number of property names");
            }

            return prop_names;
        }

        public static bool IsValidCustomPropertyName(string name)
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

        public static void CheckValidCustomPropertyName(string name)
        {
            if (!IsValidCustomPropertyName(name))
            {
                string msg = String.Format("Invalid Property Name: \"{0}\"", name);
                throw new VA.AutomationException(msg);
            }
        }

        public static bool HasCustomProperty(IVisio.Shape shape, string name)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            CheckValidCustomPropertyName(name);

            string full_prop_name = GetRowName(name);

            var exists = (short)VisExistsFlags.visExistsAnywhere;
            return 0 != (shape.CellExistsU[full_prop_name, exists]);
        }

        public static string GetRowName(string name)
        {
            return String.Format("Prop.{0}", name);
        }

        public static void DeleteCustomProperty(IVisio.Shape shape, string name)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            CheckValidCustomPropertyName(name);

            string full_prop_name = GetRowName(name);

            short row = shape.CellsU[full_prop_name].Row;
            shape.DeleteRow(CustomPropertyCells.custprop_query.Section, row);
        }

        public static void SetCustomProperty(IVisio.Shape shape, string name, string val)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            CheckValidCustomPropertyName(name);

            if (val == null)
            {
                throw new System.ArgumentNullException("val");
            }

            // create a new property
            var cp = new CustomPropertyCells();
            cp.Value = val;
            cp.Type = (int)VA.CustomProperties.FormatShapeData.String;

            SetCustomProperty(shape, name, cp);
        }
    }
}