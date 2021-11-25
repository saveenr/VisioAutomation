namespace VisioAutomation_Tests
{
    public static class TestExtensions
    {
        public static VisioAutomation.Geometry.Point GetPinPosResult(this VisioAutomation.Shapes.ShapeXFormCells xform)
        {
            return  ToPoint(xform.PinX.Value, xform.PinY.Value);
        }

        public static VisioAutomation.Geometry.Point ToPoint(string x,string y)
        {
            return new VisioAutomation.Geometry.Point(InchesToDouble(x), InchesToDouble(y));
        }

        public static double InchesToDouble(string s)
        {

            string suffix = " in.";
            string s2 = s.Substring(0, s.Length - suffix.Length);
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            return double.Parse(s2, culture);
        }

        public static void AddParameter(this System.Management.Automation.Runspaces.Command cmd, string name, object value)
        {
            var parameter = new System.Management.Automation.Runspaces.CommandParameter(name, value);
            cmd.Parameters.Add(parameter);
        }
    }
}