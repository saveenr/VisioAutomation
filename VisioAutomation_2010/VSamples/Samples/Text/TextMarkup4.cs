namespace VSamples.Samples.Text
{
    public  class TextMarkup4 : SampleMethodBase
    {
        public override void Run()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);
            var e1 = new VisioAutomation.Models.Text.Element();
            e1.AddText("This shape is ");
            e1.AddField(VisioAutomation.Models.Text.FieldConstants.Width);
            e1.AddText("inches wide by ");
            e1.AddField(VisioAutomation.Models.Text.FieldConstants.Height);
            e1.AddText("inches tall.");
            e1.SetText(s1);
        }
    }
}