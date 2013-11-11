namespace VisioPS.Commands
{
    public class RectangleCmdlet : VisioPS.VisioPSCmdlet
    {
        [System.Management.Automation.Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [System.Management.Automation.Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [System.Management.Automation.Parameter(Position = 2, Mandatory = true)]
        public double X1 { get; set; }

        [System.Management.Automation.Parameter(Position = 3, Mandatory = true)]
        public double Y1 { get; set; }

        protected VisioAutomation.Drawing.Rectangle GetRectangle()
        {
            return new VisioAutomation.Drawing.Rectangle(X0, Y0, X1, Y1);
        }
    }
}