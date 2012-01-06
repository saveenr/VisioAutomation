using VisioAutomation.Drawing;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;

namespace VisioAutomation.Text.Markup
{
    public class CharacterFormat
    {
        private double? _font_size;
        private int? _transparency;

        public CharStyle? CharStyle { get; set; }
        public int? FontID { get; set; }
        public ColorRGB? Color { get; set; }

        public CharacterFormat()
        {
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

        public void UpdateFrom(CharacterFormat other)
        {
            this.FontID = other.FontID;
            this._font_size = other._font_size;
            this.CharStyle = other.CharStyle;
            this.Color = other.Color;
            this._transparency = other._transparency;
        }

        public CharacterFormat Duplicate()
        {
            var fmt = new CharacterFormat();
            fmt.UpdateFrom(this);
            return fmt;
        }
    }
}