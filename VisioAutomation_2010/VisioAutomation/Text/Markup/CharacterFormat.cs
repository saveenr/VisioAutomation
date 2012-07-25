using VisioAutomation.Drawing;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;

namespace VisioAutomation.Text.Markup
{
    public class CharacterFormat
    {
        public ColorRGB? Color { get; set; }
        public int? Font { get; set; }
        private double? _size;
        public CharStyle? Style { get; set; }
        private int? _transparency;

        public CharacterFormat()
        {
        }

        public double? Size
        {
            get { return _size; }
            set
            {
                if (value.HasValue)
                {
                    if (value.Value <= 0)
                    {
                        throw new System.ArgumentOutOfRangeException();
                    }
                }

                _size = value;
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
            this.Font = other.Font;
            this._size = other._size;
            this.Style = other.Style;
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