using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public struct ColorHSL
    {
        // HSL http://msdn.microsoft.com/en-us/library/ms406705(v=office.12).aspx
        // HUE http://msdn.microsoft.com/en-us/library/ms406706(v=office.12).aspx
        // SAT http://msdn.microsoft.com/en-us/library/ms425560(office.12).aspx
        // LUM http://office.microsoft.com/en-us/visio-help/HV080400509.aspx

        private readonly byte _h;
        private readonly byte _s;
        private readonly byte _l;

        public ColorHSL(byte h, byte s, byte l)
        {
            _h = h;
            _s = s;
            _l = l;
        }

        private void CheckValidVisioHSL()
        {
            if (this.H > 255)
            {
                throw new System.ArgumentOutOfRangeException("h", "h must be <=255");
            }
            if (this.S > 240)
            {
                throw new System.ArgumentOutOfRangeException("s", "s must be <=240");
            }
            if (this.L > 240)
            {
                throw new System.ArgumentOutOfRangeException("l", "l must be <=240");
            }
        }

        public ColorHSL(short h, short s, short l) :
            this((byte)h, (byte)s, (byte)l)
        {
        }

        public byte H
        {
            get { return _h; }
        }

        public byte S
        {
            get { return _s; }
        }

        public byte L
        {
            get { return _l; }
        }

        public override string ToString()
        {
            var s = string.Format(System.Globalization.CultureInfo.InvariantCulture, "HSL({0},{1},{2})",this._h, this._s, this._l);
            return s;
        }

        public override bool Equals(object other)
        {
            return other is VA.Drawing.ColorHSL && Equals((VA.Drawing.ColorHSL)other);
        }

        public static bool operator ==(ColorHSL lhs, ColorHSL rhs)
        {
            return lhs.Equals(rhs);
        }

        public static bool operator !=(ColorHSL lhs, ColorHSL rhs)
        {
            return !lhs.Equals(rhs);
        }

        private bool Equals(ColorHSL other)
        {
            return (this._h == other._h && this._s == other._s && this._l == other._l);
        }

        public override int GetHashCode()
        {
            return ToHSLBytes();
        }

        /// <summary>
        /// Returns an int containing RGB values.
        /// </summary>
        /// <returns></returns>
        private int ToHSLBytes()
        {
            return (this._h << 16) | (this._s << 8) | (this._l);
        }

        public string ToFormula()
        {
            this.CheckValidVisioHSL();
            string formula = System.String.Format("HSL({0},{1},{2})", this.H, this.S, this.L);
            return formula;
        }
    }
}