using System;
using VA=VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.CustomProperties
{
    public static class CustomPropertyHelper
    {
        public static void UpdateCustomProperty(IVisio.Shape shape, string name, string val)
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
                shape.CellsSRC[(short)IVisio.VisSectionIndices.visSectionProp, row, (short)IVisio.VisCellIndices.visCustPropsValue];
            cell_propval.FormulaU = val;
        }


        public static void SetCustomProperty(
            IVisio.Shape shape,
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
                (short)IVisio.VisSectionIndices.visSectionProp,
                name,
                (short)IVisio.VisRowIndices.visRowProp);

            SetCustomProperty(shape, row, cp);
        }

        public static void SetCustomProperty(IVisio.Shape shape, short row, VA.CustomProperties.CustomPropertyCells cp)
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
        public static IDictionary<string, CustomPropertyCells> GetCustomProperties(IVisio.Shape shape)
        {
            var prop_names = GetCustomPropertyNames(shape);
            var dic = new Dictionary<string, CustomPropertyCells>(prop_names.Count);
            var cells = CustomPropertyCells.GetCells(shape);

            for (int prop_index = 0; prop_index < prop_names.Count(); prop_index++)
            {
                string prop_name = prop_names[prop_index];
                dic[prop_name] = cells[prop_index];
            }

            return dic;
        }

        public static IList<Dictionary<string, CustomPropertyCells>> GetCustomProperties(IVisio.Page page, IList<IVisio.Shape> shapes)
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
            var customprops = new List<Dictionary<string, CustomPropertyCells>>(shapeids.Count);
            var cells_list = CustomPropertyCells.GetCells(page, shapeids);
            
            if (cells_list.Count!=shapeids.Count)
            {
                throw new VA.AutomationException("1");
            }

            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shape = shapes[shape_index];
                var cells = cells_list[shape_index];
                var prop_names = GetCustomPropertyNames(shape);

                if (cells.Count != prop_names.Count)
                {
                    throw new VA.AutomationException("2");
                }

                var dic = new Dictionary<string, CustomPropertyCells>(prop_names.Count);
                
                for (int prop_index=0; prop_index< prop_names.Count(); prop_index++)
                {
                    string prop_name = prop_names[prop_index];
                    dic[prop_name] = cells[prop_index];
                }

                customprops.Add(dic);
            }

            return customprops;
        }

        public static int GetCustomPropertyCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            // If the Custom Property section does not exist then return zero immediately
            if (0 == shape.SectionExists[(short)IVisio.VisSectionIndices.visSectionProp, (short)IVisio.VisExistsFlags.visExistsAnywhere])
            {
                return 0;
            }

            var section = shape.Section[(short)IVisio.VisSectionIndices.visSectionProp];

            if (section == null)
            {
                throw new AutomationException("section is null");
            }

            int row_count = section.Shape.RowCount[(short)IVisio.VisSectionIndices.visSectionProp];

            return row_count;
        }

        public static IList<string> GetCustomPropertyNames(IVisio.Shape shape)
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
            var prop_section = shape.Section[(short)IVisio.VisSectionIndices.visSectionProp];
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

            var exists = (short)IVisio.VisExistsFlags.visExistsAnywhere;
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
            shape.DeleteRow((short)IVisio.VisSectionIndices.visSectionProp, row);
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
            cp.Type = (int)VA.CustomProperties.Format.String;

            SetCustomProperty(shape, name, cp);
        }
    }
}