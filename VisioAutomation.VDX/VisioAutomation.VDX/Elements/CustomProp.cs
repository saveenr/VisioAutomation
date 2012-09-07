using VisioAutomation.VDX.Enums;
using VA=VisioAutomation;
using VisioAutomation.VDX.Internal.Extensions;

namespace VisioAutomation.VDX.Elements
{
    public class CustomProp
    {
        // http://msdn.microsoft.com/en-us/library/aa722335.aspx

        public string NameU { get; set; }
        private int? _del;

        public string Value { get; set; }
        public VA.VDX.ShapeSheet.StringCell Prompt = new VA.VDX.ShapeSheet.StringCell();
        public VA.VDX.ShapeSheet.StringCell Label = new VA.VDX.ShapeSheet.StringCell();
        public VA.VDX.ShapeSheet.StringCell Format = new VA.VDX.ShapeSheet.StringCell();
        public VA.VDX.ShapeSheet.IntCell SortKey = new VA.VDX.ShapeSheet.IntCell();
        public VA.VDX.ShapeSheet.EnumCell<CustomPropType> Type = new VA.VDX.ShapeSheet.EnumCell<CustomPropType>(v => (int)v);
        public VA.VDX.ShapeSheet.BoolCell Invisible = new VA.VDX.ShapeSheet.BoolCell();
        public VA.VDX.ShapeSheet.BoolCell Verify = new VA.VDX.ShapeSheet.BoolCell();
        public VA.VDX.ShapeSheet.IntCell LangID = new VA.VDX.ShapeSheet.IntCell();
        public VA.VDX.ShapeSheet.EnumCell<CalendarType> Calendar = new VA.VDX.ShapeSheet.EnumCell<CalendarType>(v => (int)v);

        public string Name { get; set; }

        public int? Del
        {
            get { return _del; }
            set { _del = value; }
        }

        public int ID { get; private set; }

        public CustomProp()
        {
            
        }

        public CustomProp(int id, string nameu) :
            this()
        {
            this.ID = id;
            this.NameU = nameu;
        }

        public void AddToElement(System.Xml.Linq.XElement parent)
        {
            var prop_el = VA.VDX.Internal.XMLUtil.CreateVisioSchema2003Element("Prop");
            if (this.Name != null)
            {
                prop_el.SetElementValue("Name", this.Name);
            }

            prop_el.SetAttributeValue("NameU", this.NameU);
            prop_el.SetAttributeValue("ID", this.ID);
            prop_el.SetElementValueConditional("Del", this._del);

            if (this.Value!=null)
            {
                var val_el = new System.Xml.Linq.XElement(VA.VDX.Internal.Constants.VisioXmlNamespace2003 + "Value");
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