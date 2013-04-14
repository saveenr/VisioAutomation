using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;


namespace VisioAutomation.VDX.Sections
{
    public class Event
    {
        public DoubleCell TheData = new DoubleCell();
        public DoubleCell TheText = new DoubleCell();
        public DoubleCell EventDblClick = new DoubleCell();
        public DoubleCell EventXFMod = new DoubleCell();
        public DoubleCell EventDrop = new DoubleCell();

        public DoubleCell EventMultiDrop = new DoubleCell();

        public void AddToElement(SXL.XElement parent)
        {
            var el1 = XMLUtil.CreateVisioSchema2003Element("Event");
            el1.Add(this.TheData.ToXml("TheData"));
            el1.Add(this.TheText.ToXml("TheText"));
            el1.Add(this.EventDblClick.ToXml("EventDblClick"));
            el1.Add(this.EventXFMod.ToXml("EventXFMod"));
            el1.Add(this.EventDrop.ToXml("EventDrop"));
            parent.Add(el1);

            var el2 = XMLUtil.CreateVisioSchema2006Element("Event");
            el2.Add(this.EventMultiDrop.ToXml2006("EventMultiDrop"));
            parent.Add(el2);
        }
    }
}