using VisioAutomation.VDX.Enums;
using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections
{
    public class Char
    {
        public IntCell Font = new IntCell();
        public IntCell Color = new IntCell();
        public EnumCell<CharStyle> Style = new EnumCell<CharStyle>(v => (int) v);
        public EnumCell<CharCase> Case = new EnumCell<CharCase>(v => (int) v);
        public IntCell Pos = new IntCell();
        public DoubleCell FontScale = new DoubleCell();
        public PointCell Size = new PointCell();
        public BoolCell DoubleUnderline = new BoolCell();
        public BoolCell Overline = new BoolCell();
        public BoolCell Strikethru = new BoolCell();
        public IntCell Highlight = new IntCell();
        public BoolCell DoubleStrikethrough = new BoolCell();
        public BoolCell RTLText = new BoolCell();
        public BoolCell UseVertical = new BoolCell();
        public DoubleCell Letterspace = new DoubleCell();
        public TransparencyCell Transparency = new TransparencyCell();
        public IntCell AsianFont = new IntCell();
        public IntCell ComplexScriptFont = new IntCell();
        public IntCell LocalizeFont = new IntCell();
        public IntCell ComplexScriptSize = new IntCell();
        public IntCell LangID = new IntCell();

        public void AddToElement(SXL.XElement parent, int index)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("Char");
            el.SetAttributeValue("IX", index.ToString(System.Globalization.CultureInfo.InvariantCulture));
            el.Add(this.Font.ToXml("Font"));
            el.Add(this.Color.ToXml("Color"));

            el.Add(this.Style.ToXml("Style"));
            el.Add(this.Case.ToXml("Case"));
            el.Add(this.Pos.ToXml("Pos"));

            el.Add(this.FontScale.ToXml("FontScale"));
            el.Add(this.Size.ToXml("Size"));

            el.Add(this.DoubleUnderline.ToXml("DblUnderline"));
            el.Add(this.Overline.ToXml("Overline"));

            el.Add(this.Strikethru.ToXml("Strikethru"));
            el.Add(this.Highlight.ToXml("Highlight"));

            el.Add(this.DoubleStrikethrough.ToXml("DoubleStrikethrough"));
            el.Add(this.RTLText.ToXml("RTLText"));
            el.Add(this.UseVertical.ToXml("UseVertical"));

            el.Add(this.Letterspace.ToXml("Letterspace"));
            el.Add(this.Transparency.ToXml("ColorTrans"));

            el.Add(this.AsianFont.ToXml("AsianFont"));
            el.Add(this.ComplexScriptFont.ToXml("ComplexScriptFont"));

            el.Add(this.LocalizeFont.ToXml("LocalizeFont"));
            el.Add(this.ComplexScriptSize.ToXml("ComplexScriptSize"));
            el.Add(this.LangID.ToXml("LangID"));

            parent.Add(el);
        }
    }
}