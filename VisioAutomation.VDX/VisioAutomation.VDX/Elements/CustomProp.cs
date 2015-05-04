using VisioAutomation.VDX.Enums;
using VA=VisioAutomation;
using VisioAutomation.VDX.Internal.Extensions;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Elements
{
    public class CustomProp
    {
        // http://msdn.microsoft.com/en-us/library/aa722335.aspx

        public string NameU { get; set; }
        private int? _del;

        public string Value { get; set; }
        public ShapeSheet.StringCell Prompt = new ShapeSheet.StringCell();
        public ShapeSheet.StringCell Label = new ShapeSheet.StringCell();
        public ShapeSheet.StringCell Format = new ShapeSheet.StringCell();
        public ShapeSheet.IntCell SortKey = new ShapeSheet.IntCell();
        public ShapeSheet.EnumCell<CustomPropType> Type = new ShapeSheet.EnumCell<CustomPropType>(v => (int)v);
        public ShapeSheet.BoolCell Invisible = new ShapeSheet.BoolCell();
        public ShapeSheet.BoolCell Verify = new ShapeSheet.BoolCell();
        public ShapeSheet.IntCell LangID = new ShapeSheet.IntCell();
        public ShapeSheet.EnumCell<CalendarType> Calendar = new ShapeSheet.EnumCell<CalendarType>(v => (int)v);

        public string Name { get; set; }

        public int? Del
        {
            get { return this._del; }
            set { this._del = value; }
        }

        public int ID { get; internal set; }

        public CustomProp()
        {
            
        }

        public CustomProp(string nameu) :
            this()
        {
            this.ID = -1;
            this.NameU = nameu;
        }

        public void AddToElement(SXL.XElement parent)
        {
            var prop_el = Internal.XMLUtil.CreateVisioSchema2003Element("Prop");
            if (this.Name != null)
            {
                prop_el.SetElementValue("Name", this.Name);
            }

            prop_el.SetAttributeValue("NameU", this.NameU);
            prop_el.SetAttributeValue("ID", this.ID);
            prop_el.SetElementValueConditional("Del", this._del);

            if (this.Value!=null)
            {
                var val_el = new SXL.XElement(Internal.Constants.VisioXmlNamespace2003 + "Value");
                prop_el.Add(val_el);
                val_el.SetAttributeValue("Unit", "STR");
                val_el.SetValue(this.Value);                
            }
            prop_el.Add(this.Prompt.ToXml("Prompt"));
            prop_el.Add(this.Label.ToXml("Label"));
            prop_el.Add(this.Format.ToXml("Format"));
            prop_el.Add(this.SortKey.ToXml("SortKey"));
            prop_el.Add(this.Type.ToXml("Type"));
            prop_el.Add(this.Invisible.ToXml("Invisible"));
            prop_el.Add(this.LangID.ToXml("LangID"));
            prop_el.Add(this.Calendar.ToXml("Calendar"));
            parent.Add(prop_el);
        }
    }
}