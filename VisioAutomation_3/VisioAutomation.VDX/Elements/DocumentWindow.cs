using VA=VisioAutomation;
using VisioAutomation.VDX.Internal.Extensions;

namespace VisioAutomation.VDX.Elements
{
    public class DocumentWindow : Window
    {
        public DocumentWindow() :
            base()
        {
        }

        public int Page { get; set; }

        public double ViewScale { get; set; }

        public string Document { get; set; }

        public bool? ShowRulers { get; set; }

        public bool? ShowGrid { get; set; }

        public bool? ShowPageBreaks { get; set; }

        public bool? ShowGuides { get; set; }

        public bool? ShowConnectionPoints { get; set; }

        public VA.VDX.Enums.GlueType? GlueSettings { get; set; }

        public SnapType? SnapSettings { get; set; }

        public VA.VDX.Enums.SnapExtensionsType? SnapExtensions { get; set; }

        public int DynamicGridEnabled { get; set; }

        public double? TabSplitterPos { get; set; }


        public override void AddToElement(System.Xml.Linq.XElement parent)
        {
            string ns_2003 = VA.VDX.Internal.Constants.VisioXmlNamespace2003;

            var window_el = VA.VDX.Internal.XMLUtil.CreateVisioSchema2003Element("Window");
            window_el.SetAttributeValue("ID", this.ID);
            window_el.SetAttributeValue("WindowType", "Drawing");
            window_el.SetAttributeValue("ContainerType", "Page");
            window_el.SetAttributeValue("Page", this.Page);
            window_el.SetElementValueConditionalBool(ns_2003 + "ShowRulers", this.ShowRulers);
            window_el.SetElementValueConditionalBool(ns_2003 + "ShowGrid", this.ShowGrid);
            window_el.SetElementValueConditionalBool(ns_2003 + "ShowPageBreaks",this.ShowPageBreaks);
            window_el.SetElementValueConditionalBool(ns_2003 + "ShowGuides", this.ShowGuides);
            window_el.SetElementValueConditionalBool(ns_2003 + "ShowConnectionPoints", this.ShowConnectionPoints);
            window_el.SetElementValueConditional(ns_2003 + "GlueSettings", this.GlueSettings, v => (int) v);
            window_el.SetElementValueConditional(ns_2003 + "SnapSettings", this.SnapSettings, v => (int) v);
            window_el.SetElementValueConditional(ns_2003 + "SnapExtensions", this.SnapExtensions, v => (int) v);
            parent.Add(window_el);
        }
    }
}