using System.Xml.Linq;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;
using System.Collections.Generic;

namespace VisioAutomation.Text.Markup
{
    public class TextFormat
    {
        private double? _font_size;
        private int? _transparency;

        public AlignmentHorizontal? HAlign { get; set; }
        public ColorRGB? Color { get; set; }
        public double? Indent { get; set; }
        public bool? Bullets { get; set; }
        public CharStyle? CharStyle { get; set; }

        public TextFormat()
        {
        }

        public string Font { get; set; }

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

        public void UpdateFrom(TextFormat other)
        {
            this.Font = other.Font;
            this._font_size = other._font_size;
            this.HAlign=other.HAlign;
            this.CharStyle = other.CharStyle;
            this.Color = other.Color;
            this.Indent =other.Indent;
            this.Bullets= other.Bullets;
            this._transparency = other._transparency;
        }

        public TextFormat Duplicate()
        {
            var fmt = new TextFormat();
            fmt.UpdateFrom(this);
            return fmt;
        }

        public override string ToString()
        {
            string[] a =
                {
                    GetNameValuePairString(this.Font, "Font"),
                    GetNameValuePairString(this._font_size, "FontSize"),
                    GetNameValuePairString(this.HAlign, "HAlign"),
                    GetNameValuePairString(this.CharStyle, "CharStyle"),
                    GetNameValuePairString(this.Color, "Color"),
                    GetNameValuePairString(this.Indent, "Indent"),
                    GetNameValuePairString(this.Bullets, "Bullets"),
                    GetNameValuePairString(this._transparency, "Transparency")
                };
            string s = string.Join(",", a);
            return s;
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