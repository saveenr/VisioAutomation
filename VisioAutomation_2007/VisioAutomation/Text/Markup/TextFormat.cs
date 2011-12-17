using System.Xml.Linq;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;
using System.Collections.Generic;

namespace VisioAutomation.Text.Markup
{
    public class TextFormat
    {
        private string _font;
        private double? _font_size;
        private VA.Drawing.AlignmentHorizontal? _h_align;
        private VA.Drawing.ColorRGB? _color;
        private double? _indent;
        private bool? _bullets;
        private int? _transparency;
        private VA.Text.CharStyle? _char_style;

        public TextFormat()
        {
            this.initialize_markup_attributes();
        }

        public string Font
        {
            get { return _font; }
            set
            {
                _font = value;
            }
        }

        public double? FontSize
        {
            get { return _font_size; }
            set
            {
                if (value.HasValue)
                {
                    if (value.Value <= 0)
                    {
                        throw new System.ArgumentOutOfRangeException();
                    }
                }

                _font_size = value;
            }
        }

        public VA.Drawing.AlignmentHorizontal? HAlign
        {
            get { return _h_align; }
            set { _h_align = value; }
        }

        public VA.Drawing.ColorRGB? Color
        {
            get { return _color; }
            set { _color = value; }
        }

        public double? Indent
        {
            get { return _indent; }
            set { _indent = value; }
        }

        public bool? Bullets
        {
            get { return _bullets; }
            set { _bullets = value; }
        }

        public int? Transparency
        {
            get { return _transparency; }
            set
            {
                if (value.HasValue)
                {
                    if (value.Value < 0)
                    {
                        throw new System.ArgumentOutOfRangeException();
                    }
                }

                if (value.HasValue)
                {
                    if (value.Value > 100)
                    {
                        throw new System.ArgumentOutOfRangeException();
                    }
                }


                _transparency = value;
            }
        }

        public VA.Text.CharStyle? CharStyle
        {
            get { return _char_style; }
            set { _char_style = value; }
        }

        public void UpdateFrom(TextFormat other)
        {
            this._font = other._font;
            this._font_size = other._font_size;
            this._h_align=other._h_align;
            this._char_style = other._char_style;
            this._color = other._color;
            this._indent =other._indent;
            this._bullets= other._bullets;
            this._transparency = other._transparency;
        }

        public TextFormat Duplicate()
        {
            var fmt = new TextFormat();
            fmt.UpdateFrom(this);
            return fmt;
        }

        public void LoadAttributesFromXml(SXL.XElement el)
        {
            load_font(el);
            load_fontsize(el);
            load_halign(el);
            load_charstyle(el);
            load_color(el);
            load_transparency(el);
            load_indent(el);
            load_bullets(el);
        }

        private void load_bullets(XElement el)
        {
            var b_attr = el.Attribute("bullets");
            if (b_attr == null)
            {
                this._bullets = null;
            }
            else
            {
                this._bullets = string_to_bool(b_attr.Value);
            }
        }

        private void load_indent(XElement el)
        {
            var in_attr = el.Attribute("indent");
            if (in_attr == null)
            {
                this._indent=null;
            }
            else
            {
                this._indent = double.Parse(in_attr.Value);
            }
        }

        private void load_transparency(XElement el)
        {
            var tr_attr = el.Attribute("transparency");
            if (tr_attr == null)
            {
                this._transparency = null;
            }
            else
            {
                this._transparency = int.Parse(tr_attr.Value);
            }
        }

        private void load_color(XElement el)
        {
            var color_attr = el.Attribute("color");
            if (color_attr == null)
            {
                this._color=null;
            }
            else
            {
                var cf = VA.Drawing.ColorRGB.ParseWebColor(color_attr.Value);
                this._color = cf;
            }
        }

        private void load_charstyle(XElement el)
        {
            VA.Text.CharStyle cs = 0;

            if (attr_to_style == null)
            {
                attr_to_style = new Dictionary<string, VA.Text.CharStyle>
                                    {
                                        {"bold", VA.Text.CharStyle.Bold},
                                        {"italic", VA.Text.CharStyle.Italic},
                                        {"underline", VA.Text.CharStyle.UnderLine},
                                        {"smallcaps", VA.Text.CharStyle.SmallCaps},
                                    };
            }

            foreach (string attr_name in attr_to_style.Keys)
            {
                if (get_bool_attribute_value(el, attr_name))
                {
                    cs = cs | attr_to_style[attr_name];
                }
            }

            this._char_style = cs;
        }

        private void load_halign(XElement el)
        {
            var halign_attr = el.Attribute("halign");
            if (halign_attr == null)
            {
                this._h_align=null;
            }
            else
            {
                this._h_align = parseenum<VA.Drawing.AlignmentHorizontal>(halign_attr.Value, true);
            }
        }

        private void load_fontsize(XElement el)
        {
            var fontsize_attr = el.Attribute("size");
            if (fontsize_attr == null)
            {
                this._font_size=null;
            }
            else
            {
                this._font_size = double.Parse(fontsize_attr.Value);
            }
        }

        private void load_font(XElement el)
        {
            var font_attr = el.Attribute("font");
            if (font_attr == null)
            {
                this._font=null;
            }
            else
            {
                this._font = font_attr.Value;
            }
        }

        public void initialize_markup_attributes()
        {

        }

        public override string ToString()
        {
            string[] a =
                {
                    GetNameValuePairString(this._font, "Font"),
                    GetNameValuePairString(this._font_size, "FontSize"),
                    GetNameValuePairString(this._h_align, "HAlign"),
                    GetNameValuePairString(this._char_style, "CharStyle"),
                    GetNameValuePairString(this._color, "Color"),
                    GetNameValuePairString(this._indent, "Indent"),
                    GetNameValuePairString(this._bullets, "Bullets"),
                    GetNameValuePairString(this._transparency, "Transparency")
                };
            string s = string.Join(",", a);
            return s;
        }

        private static T parseenum<T>(string s, bool ignorecase)
        {
            T outval = (T) System.Enum.Parse(typeof (T), s, ignorecase);
            return outval;
        }

        private static Dictionary<string, VA.Text.CharStyle>
            attr_to_style;


        private static bool get_bool_attribute_value(SXL.XElement node, string attr_name)
        {
            var attr = node.Attribute(attr_name);
            if (attr != null)
            {
                string attr_value = attr.Value;
                return string_to_bool(attr_value);
            }

            return false;
        }

        private static bool string_to_bool(string str_value)
        {
            if (str_value == "1")
            {
                return true;
            }
            else if (str_value == "0")
            {
                return false;
            }
            else
            {
                string msg = string.Format("must be either 0 or 1");
                throw new AutomationException(msg);
            }
        }

        private static string GetNameValuePairString(string optval, string name)
        {
            string result;

            if (optval!=null)
            {
                result = string.Format("{0}=\"{1}\"", name, optval);
            }
            else
            {
                result = string.Format("{0}={1}", name, "n/a");
            }

            return result;
        }

        private static string GetNameValuePairString<T>(T? optval, string name) where T: struct
        {
            string result;

            if (optval.HasValue)
            {
                result = string.Format("{0}=\"{1}\"", name, optval.Value);
            }
            else
            {
                result = string.Format("{0}={1}", name, "n/a");
            }

            return result;
        }
    }
}