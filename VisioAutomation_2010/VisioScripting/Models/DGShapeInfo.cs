using VisioScripting.Extensions;
using VisioAutomation.Shapes;
using SXL = System.Xml.Linq;

namespace VisioScripting.Models
{
    internal class DgShapeInfo
    {
        public string ID;
        public string Label;
        public string Stencil;
        public string Master;
        public string Url;
        public SXL.XElement Element;

        public CustomPropertyDictionary CustProps;

        public static DgShapeInfo FromXml(Client client, SXL.XElement shape_el)
        {
            var info = new DgShapeInfo();
            info.ID = shape_el.Attribute("id").Value;
            client.Output.WriteVerbose( "Reading shape id={0}", info.ID);

            info.Label = shape_el.Attribute("label").Value;
            info.Stencil = shape_el.Attribute("stencil").Value;
            info.Master = shape_el.Attribute("master").Value;
            info.Element = shape_el;
            info.Url = shape_el.GetAttributeValue("url", null);

            info.CustProps = DgXml.ParseCustomProps(shape_el, info.ID);

            return info;
        }
    }
}