using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Models.Dom;
using VisioAutomation.Shapes;
using VisioScripting.Extensions;
using SXL = System.Xml.Linq;

namespace VisioScripting.Models
{
    // Parsing helpers for the optional parts of the directed graph XML format that are
    // shared between <shape> and <connector>: <customprop>, <cells>, <hyperlink> and the
    // enum / size attributes.
    internal static class DgXml
    {
        private static readonly System.Globalization.CultureInfo _culture = System.Globalization.CultureInfo.InvariantCulture;

        private static readonly Dictionary<string, System.Reflection.PropertyInfo> _cell_properties = _build_cell_property_map();

        private static Dictionary<string, System.Reflection.PropertyInfo> _build_cell_property_map()
        {
            var map = new Dictionary<string, System.Reflection.PropertyInfo>(System.StringComparer.OrdinalIgnoreCase);
            foreach (var prop in typeof(ShapeCells).GetProperties())
            {
                if (prop.PropertyType == typeof(VisioAutomation.Core.CellValue) && prop.CanWrite)
                {
                    map[prop.Name] = prop;
                }
            }
            return map;
        }

        public static double ParseDouble(string text)
        {
            return double.Parse(text, _culture);
        }

        public static VisioAutomation.Models.ConnectorType ParseConnectorType(string text)
        {
            return (VisioAutomation.Models.ConnectorType)System.Enum.Parse(
                typeof(VisioAutomation.Models.ConnectorType), text, ignoreCase: true);
        }

        // Reads a pair of optional attributes (for example "borderwidth" and "borderheight") into a Size.
        // Returns null if neither is present. If only one is present the other keeps the value from "current".
        public static VisioAutomation.Core.Size? GetOptionalSizePair(SXL.XElement el, string width_attr, string height_attr, VisioAutomation.Core.Size current)
        {
            var w = el.Attribute(width_attr);
            var h = el.Attribute(height_attr);
            if (w == null && h == null)
            {
                return null;
            }

            double width = w != null ? ParseDouble(w.Value) : current.Width;
            double height = h != null ? ParseDouble(h.Value) : current.Height;
            return new VisioAutomation.Core.Size(width, height);
        }

        // A shape's "width" and "height" attributes. Both or neither: the size cannot be completed from the master here.
        public static VisioAutomation.Core.Size? GetShapeSize(SXL.XElement shape_el, string id)
        {
            var w = shape_el.Attribute("width");
            var h = shape_el.Attribute("height");
            if (w == null && h == null)
            {
                return null;
            }

            if (w == null || h == null)
            {
                string msg = string.Format(_culture, "Shape \"{0}\" must set both \"width\" and \"height\", or neither.", id);
                throw new System.ArgumentException(msg);
            }

            return new VisioAutomation.Core.Size(ParseDouble(w.Value), ParseDouble(h.Value));
        }

        public static CustomPropertyDictionary ParseCustomProps(SXL.XElement parent, string id)
        {
            var dic = new CustomPropertyDictionary();
            foreach (var customprop_el in parent.Elements("customprop"))
            {
                string cp_name = customprop_el.Attribute("name").Value;
                string cp_value = customprop_el.Attribute("value").Value;
                string cp_type = customprop_el.GetAttributeValue("type", "string");

                var cp = new CustomPropertyCells();
                if (string.Equals(cp_type, "string", System.StringComparison.OrdinalIgnoreCase))
                {
                    cp.SetString(cp_value);
                }
                else if (string.Equals(cp_type, "number", System.StringComparison.OrdinalIgnoreCase))
                {
                    cp.SetNumber(ParseDouble(cp_value));
                }
                else if (string.Equals(cp_type, "boolean", System.StringComparison.OrdinalIgnoreCase)
                    || string.Equals(cp_type, "bool", System.StringComparison.OrdinalIgnoreCase))
                {
                    cp.SetBool(bool.Parse(cp_value));
                }
                else if (string.Equals(cp_type, "date", System.StringComparison.OrdinalIgnoreCase))
                {
                    cp.SetDate(System.DateTime.Parse(cp_value, _culture));
                }
                else
                {
                    string msg = string.Format(_culture,
                        "\"{0}\" : custom property \"{1}\" has unsupported type \"{2}\". Use string, number, boolean or date.",
                        id, cp_name, cp_type);
                    throw new System.ArgumentException(msg);
                }

                var label = customprop_el.Attribute("label");
                if (label != null)
                {
                    cp.Label = label.Value;
                }

                var prompt = customprop_el.Attribute("prompt");
                if (prompt != null)
                {
                    cp.Prompt = prompt.Value;
                }

                var format = customprop_el.Attribute("format");
                if (format != null)
                {
                    cp.Format = format.Value;
                }

                dic.Add(cp_name, cp);
            }

            return dic;
        }

        // Sets cells from a <cells><cell name="..." value="..." /></cells> child. The cell names are the property
        // names of VisioAutomation.Models.Dom.ShapeCells (for example FillForeground or LineColor), matched without
        // regard to case, and the value is a Visio formula or literal. Returns true if a <cells> element was present.
        public static bool ApplyCells(SXL.XElement parent, ShapeCells cells, string id)
        {
            var cells_el = parent.Element("cells");
            if (cells_el == null)
            {
                return false;
            }

            foreach (var cell_el in cells_el.Elements("cell"))
            {
                string name = cell_el.Attribute("name").Value;
                string value = cell_el.Attribute("value").Value;

                System.Reflection.PropertyInfo prop;
                if (!_cell_properties.TryGetValue(name, out prop))
                {
                    string msg = string.Format(_culture,
                        "\"{0}\" : unknown cell \"{1}\". Cell names are the property names of VisioAutomation.Models.Dom.ShapeCells, for example FillForeground or LineColor.",
                        id, name);
                    throw new System.ArgumentException(msg);
                }

                prop.SetValue(cells, new VisioAutomation.Core.CellValue(value));
            }

            return true;
        }

        public static List<Hyperlink> ParseHyperlinks(SXL.XElement shape_el)
        {
            var list = new List<Hyperlink>();
            foreach (var hl_el in shape_el.Elements("hyperlink"))
            {
                var hl = new Hyperlink(hl_el.Attribute("name").Value, hl_el.Attribute("address").Value);
                hl.SubAddress = hl_el.GetAttributeValue("subaddress", null);
                hl.Description = hl_el.GetAttributeValue("description", null);
                list.Add(hl);
            }
            return list;
        }
    }
}
