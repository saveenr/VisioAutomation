using System.Xml.Linq;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;
using System.Collections.Generic;

namespace VisioAutomation.Text.Markup
{
    public class TextParaFormat
    {
        private bool? _bullets;
        private double? _indent;
        private VA.Drawing.AlignmentHorizontal? _h_align;

        public VA.Drawing.AlignmentHorizontal? HAlign
        {
            get { return _h_align; }
            set { _h_align = value; }
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

        public void UpdateFrom(TextParaFormat other)
        {
            this.HAlign =  other.HAlign;
            this.Indent =  other.Indent;
            this.Bullets = other.Bullets;
        }

        public TextParaFormat Duplicate()
        {
            var fmt = new TextParaFormat();
            fmt.UpdateFrom(this);
            return fmt;
        }
    }

    public class TextCharFormat
    {
        private string _font;
        private double? _font_size;
        private VA.Drawing.ColorRGB? _color;
        private int? _transparency;
        private VA.Text.CharStyle? _char_style;

        public TextCharFormat()
        {
        }

        public string Font
        {
            get { return _font; }
            set
            {
                _font = value;
            }
        }

        public VA.Drawing.ColorRGB? Color
        {
            get { return _color; }
            set { _color = value; }
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

        public void UpdateFrom(TextCharFormat other)
        {
            this._font = other._font;
            this._font_size = other._font_size;
            this._char_style = other._char_style;
            this._color = other._color;
            this._transparency = other._transparency;
        }

        public TextCharFormat Duplicate()
        {
            var fmt = new TextCharFormat();
            fmt.UpdateFrom(this);
            return fmt;
        }
    }
}