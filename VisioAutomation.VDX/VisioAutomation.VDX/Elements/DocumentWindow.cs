using VA=VisioAutomation;
using VisioAutomation.VDX.Internal.Extensions;
using SXL = System.Xml.Linq;
using System.Linq;

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

        public Enums.GlueType? GlueSettings { get; set; }

        public SnapType? SnapSettings { get; set; }

        public Enums.SnapExtensionsType? SnapExtensions { get; set; }

        public int DynamicGridEnabled { get; set; }

        public double? TabSplitterPos { get; set; }

        internal void ValidatePage( Drawing drawing )
        {
            if (drawing==null)
            {
                throw new System.ArgumentNullException("drawing");
            }

            if (this.Page < 0)
            {
                throw new System.ArgumentException("Negative page not allowed in document window");
            }

            bool found = drawing.Pages.Items.Any(p => this.Page == p.ID);

            if (!found)
            {
                string msg = "Document window pointing to page that does not exist";
                throw new System.ArgumentException(msg);
            }
        }
        public override void AddToElement(SXL.XElement parent)
        {
            string ns_2003 = Internal.Constants.VisioXmlNamespace2003;

            var window_el = Internal.XMLUtil.CreateVisioSchema2003Element("Window");
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