using VA = VisioAutomation;

namespace VisioAutomation.Drawing
{
    [System.Diagnostics.DebuggerDisplay("{Value}")]
    public struct Transparency
    {
        public readonly double Value;

        public Transparency(double v)
        {
            if ((v < 0) || (v > 1.0))
            {
                throw new System.ArgumentOutOfRangeException();
            }
            Value = v;
        }


        public static implicit operator Transparency(double v)
        {
            return new Transparency(v);
        }

        public string ToFormula()
        {
            var invariantCulture = System.Globalization.CultureInfo.InvariantCulture;
            var formula = this.Value.ToString(invariantCulture);
            return formula;
        }
    }
}