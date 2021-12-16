namespace VSamples.Samples.Text
{
    public  class TextFields : SampleMethod
    {

        public override void Run()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var s0 = page.DrawRectangle(1, 1, 4, 4);

            var markup1 = new VisioAutomation.Models.Text.Element();
            markup1.AddField(VisioAutomation.Models.Text.FieldConstants.PageName);
            markup1.AddText(" (");
            markup1.AddField(VisioAutomation.Models.Text.FieldConstants.PageNumber);
            markup1.AddText(" of ");
            markup1.AddField(VisioAutomation.Models.Text.FieldConstants.NumberOfPages);
            markup1.AddText(") ");
            markup1.SetText(s0);
        }
    }
}