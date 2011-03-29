using System.Collections.Generic;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.FlowChart
{
    internal class ShapeInfo
    {
        public string ID;
        public string Label;
        public string Stencil;
        public string Master;
        public string URL;
        public System.Xml.Linq.XElement Element;

        public Dictionary<string, VA.CustomProperties.CustomPropertyCells> custprops;

        public static ShapeInfo FromXml(Session scriptingsession, System.Xml.Linq.XElement shape_el)
        {
            var info = new ShapeInfo();
            info.ID = shape_el.Attribute("id").Value;
            scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Reading shape id={0}", info.ID);

            info.Label = shape_el.Attribute("label").Value;
            info.Stencil = shape_el.Attribute("stencil").Value;
            info.Master = shape_el.Attribute("master").Value;
            info.Element = shape_el;
            info.URL = VA.Scripting.XmlUtil.GetAttributeValue(shape_el, "url", null);

            info.custprops = new Dictionary<string, VA.CustomProperties.CustomPropertyCells>();
            foreach (var customprop_el in shape_el.Elements("customprop"))
            {
                string cp_name = customprop_el.Attribute("name").Value;
                string cp_value = customprop_el.Attribute("value").Value;

                var cp = new VA.CustomProperties.CustomPropertyCells();
                cp.Value = cp_value;

                info.custprops.Add(cp_name,cp);
            }

            return info;
        }
    }
}