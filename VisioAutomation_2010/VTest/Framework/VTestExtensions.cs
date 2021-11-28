namespace VTest.Framework
{
    public static class VTestExtensions
    {
        public static VisioAutomation.Core.Point GetPinPosResult(this VisioAutomation.Shapes.XFormCells xform)
        {
            return  ToPoint(xform.PinX.Value, xform.PinY.Value);
        }

        public static VisioAutomation.Core.Point ToPoint(string x,string y)
        {
            return new VisioAutomation.Core.Point(InchesToDouble(x), InchesToDouble(y));
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