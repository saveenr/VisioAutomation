using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections
{
    public class EventX
    {
        public DoubleCell EventMultiDrop = new DoubleCell();

        public void AddToElement(SXL.XElement parent)
        {
            var el = XMLUtil.CreateVisioSchema2006Element("Event");
            el.Add(this.EventMultiDrop.ToXml2006("EventMultiDrop"));
            parent.Add(el);
        }
    }
}