namespace VisioAutomation_Tests
{
    public static class TestExtensions
    {
        public static VisioAutomation.Drawing.Point GetPinPosResult(this VisioAutomation.Shapes.XFormCells xform)
        {
            return new VisioAutomation.Drawing.Point(xform.PinX.Result, xform.PinY.Result);
        }

        public static void AddParameter(this System.Management.Automation.Runspaces.Command cmd, string name, object value)
        {
            var parameter = new System.Management.Automation.Runspaces.CommandParameter(name, value);
            cmd.Parameters.Add(parameter);
        }
    }
}