using VisioAutomation.VDX.Internal;

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

        public void ToXml(System.Xml.Linq.XElement parent)
        {
            var facename_el = XMLUtil.CreateVisioSchema2003Element("FaceName");
            facename_el.SetAttributeValue("ID", this.ID.ToString(System.Globalization.CultureInfo.InvariantCulture));
            facename_el.SetAttributeValue("Name", this.Name);
            parent.Add(facename_el);
        }
    }
}