using SXL = System.Xml.Linq;

namespace VisioScripting.Models
{
    internal class DgConnectorInfo
    {
        public string ID;
        public string Label;
        public string From;
        public string To;
        public SXL.XElement Element;

        public static DgConnectorInfo FromXml(Client client, SXL.XElement shape_el)
        {
            var info = new DgConnectorInfo();
            info.ID = shape_el.Attribute("id").Value;
            client.Output.WriteVerbose("Reading connector id={0}", info.ID);

            info.Label = shape_el.Attribute("label").Value;
            info.From = shape_el.Attribute("from").Value;
            info.To = shape_el.Attribute("to").Value;

            info.Element = shape_el;
            return info;
        }
    }
}