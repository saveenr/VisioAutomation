namespace VisioPS.Commands
{
    public class RectangleCmdlet : VisioPS.VisioPSCmdlet
    {
        [System.Management.Automation.Parameter(Position = 0, Mandatory = true, ParameterSetName = "NumberArray")]
        public double[] Doubles { get; set; }

        [System.Management.Automation.Parameter(Position = 0, Mandatory = true, ParameterSetName = "SeparateNumbers")]
        public double X0 { get; set; }

        [System.Management.Automation.Parameter(Position = 1, Mandatory = true, ParameterSetName = "SeparateNumbers")]
        public double Y0 { get; set; }

        [System.Management.Automation.Parameter(Position = 2, Mandatory = true, ParameterSetName = "SeparateNumbers")]
        public double X1 { get; set; }

        [System.Management.Automation.Parameter(Position = 3, Mandatory = true, ParameterSetName = "SeparateNumbers")]
        public double Y1 { get; set; }

        protected VisioAutomation.Drawing.Rectangle GetRectangle()
        {
            if (this.Doubles != null)
            {
                if (Doubles.Length != 4)
                {
                    string msg = "Must have four vales";
                    throw new System.ArgumentOutOfRangeException(msg);
                }
                return new VisioAutomation.Drawing.Rectangle(Doubles[0], Doubles[1], Doubles[2], Doubles[3]);
            }
            else
            {
                return new VisioAutomation.Drawing.Rectangle(X0, Y0, X1, Y1);
            }
        }
    }
}