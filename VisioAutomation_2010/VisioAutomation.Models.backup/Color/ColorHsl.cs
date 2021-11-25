﻿namespace VisioAutomation.Models.Color
{

    public struct ColorHsl
    {
        // HSL http://msdn.microsoft.com/en-us/library/ms406705(v=office.12).aspx
        // HUE http://msdn.microsoft.com/en-us/library/ms406706(v=office.12).aspx
        // SAT http://msdn.microsoft.com/en-us/library/ms425560(office.12).aspx
        // LUM http://office.microsoft.com/en-us/visio-help/HV080400509.aspx

        private readonly byte _h;
        private readonly byte _s;
        private readonly byte _l;

        public ColorHsl(byte h, byte s, byte l)
        {
            this._h = h;
            this._s = s;
            this._l = l;
        }

        private void _check_valid_visio_hsl()
        {
            _check_valid_visio_hsl(this.H,this.S,this.L);
        }

        private static void _check_valid_visio_hsl(byte h, byte s, byte l)
        {
            if (h > 255)
            {
                throw new System.ArgumentOutOfRangeException(nameof(h), "Visio Hue value must be <=255");
            }
            if (s > 240)
            {
                throw new System.ArgumentOutOfRangeException(nameof(s), "Visio saturation value must be <=240");
            }
            if (l > 240)
            {
                throw new System.ArgumentOutOfRangeException(nameof(l), "Visio lumincance value must be <=240");
            }
        }

        public ColorHsl(short h, short s, short l) :
            this((byte)h, (byte)s, (byte)l)
        {
        }

        public byte H => _h;

        public byte S => _s;

        public byte L => _l;

        public override string ToString()
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            var s = string.Format(culture, "{0}({1},{2},{3})", nameof(ColorHsl), this.H, this.S, this.L);
            return s;
        }

        public override bool Equals(object other)
        {
            return other is ColorHsl && this._equals((ColorHsl)other);
        }

        public static bool operator ==(ColorHsl lhs, ColorHsl rhs)
        {
            return lhs._equals(rhs);
        }

        public static bool operator !=(ColorHsl lhs, ColorHsl rhs)
        {
            return !lhs._equals(rhs);
        }

        private bool _equals(ColorHsl other)
        {
            return (this.H == other.H && this.S == other.S && this.L == other.L);
        }

        public override int GetHashCode()
        {
            return this._to_hsl_bytes();
        }

        private int _to_hsl_bytes()
        {
            return (this.H << 16) | (this.S << 8) | (this.L);
        }

        public string ToFormula()
        {
            this._check_valid_visio_hsl();
            string formula = string.Format("{0}({1},{2},{3})", "HSL",this.H, this.S, this.L);
            return formula;
        }
    }
}