namespace VSamples.Samples.Text
{
    public  class NonRotatingText : SampleMethodBase
    {
        public override void Run()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var s0 = page.DrawRectangle(1, 1, 4, 4);
            s0.Text = "Hello World";

            var src = VisioAutomation.Core.SrcConstants.TextXFormAngle;
            var cell = s0.CellsSRC[src.Section, src.Row, src.Cell];
            cell.Formula = "-Angle";
        }
    }
}