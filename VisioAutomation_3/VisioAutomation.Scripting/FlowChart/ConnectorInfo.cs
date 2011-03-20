using VAS = VisioAutomation.Scripting;

namespace VisioAutomation.Scripting.FlowChart
{
    internal class ConnectorInfo
    {
        public string ID;
        public string Label;
        public string From;
        public string To;
        public System.Xml.Linq.XElement Element;

        public static ConnectorInfo FromXml(Session scriptingsession, System.Xml.Linq.XElement shape_el)
        {
            var info = new ConnectorInfo();
            info.ID = shape_el.Attribute("id").Value;
            scriptingsession.Write(VAS.OutputStream.Verbose,"Reading connector id={0}", info.ID);

            info.Label = shape_el.Attribute("label").Value;
            info.From = shape_el.Attribute("from").Value;
            info.To = shape_el.Attribute("to").Value;

            info.Element = shape_el;
            return info;
        }
    }
}