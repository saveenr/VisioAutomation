using VisioAutomation.VDX.Internal.Extensions;
using VisioAutomation.VDX.Internal;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Elements
{
    public class Face : Node
    {
        public int ID { get; private set; }
        public string Name { get; set; }

        public Face(int id, string name)
        {
            this.ID = id;
            this.Name = name;
        }

        public void ToXml(SXL.XElement parent)
        {
            var facename_el = XMLUtil.CreateVisioSchema2003Element("FaceName");
            facename_el.SetAttributeValueInt("ID", this.ID);
            facename_el.SetAttributeValue("Name", this.Name);
            parent.Add(facename_el);
        }
    }
}