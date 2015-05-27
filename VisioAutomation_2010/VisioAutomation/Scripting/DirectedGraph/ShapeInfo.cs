using System.Collections.Generic;
using VACUSTPROP = VisioAutomation.Shapes.CustomProperties;
using SXL = System.Xml.Linq;

namespace VisioAutomation.Scripting.DirectedGraph
{
    internal class ShapeInfo
    {
        public string ID;
        public string Label;
        public string Stencil;
        public string Master;
        public string URL;
        public SXL.XElement Element;

        public Dictionary<string, VACUSTPROP.CustomPropertyCells> custprops;

        public static ShapeInfo FromXml(Client client, SXL.XElement shape_el)
        {
            var info = new ShapeInfo();
            info.ID = shape_el.Attribute("id").Value;
            client.WriteVerbose( "Reading shape id={0}", info.ID);

            info.Label = shape_el.Attribute("label").Value;
            info.Stencil = shape_el.Attribute("stencil").Value;
            info.Master = shape_el.Attribute("master").Value;
            info.Element = shape_el;
            info.URL = XmlUtil.GetAttributeValue(shape_el, "url", null);

            info.custprops = new Dictionary<string, VACUSTPROP.CustomPropertyCells>();
            foreach (var customprop_el in shape_el.Elements("customprop"))
            {
                string cp_name = customprop_el.Attribute("name").Value;
                string cp_value = customprop_el.Attribute("value").Value;

                var cp = new VACUSTPROP.CustomPropertyCells();
                cp.Value = cp_value;

                info.custprops.Add(cp_name,cp);
            }

            return info;
        }
    }
}