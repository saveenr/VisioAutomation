namespace VSamples.Samples.Text
{
    public  class TextMarkup1 : SampleMethod
    {
        public override void Run()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);
            var tnr = page.Document.Fonts["Times New Roman"];

            var e1 = new VisioAutomation.Models.Text.Element();
            var color_red = new VisioAutomation.Models.Color.ColorRgb(0xff0000);
            e1.CharacterFormatting.Color = color_red.ToFormula();
            e1.CharacterFormatting.Font = tnr.ID;
            e1.CharacterFormatting.Font = "20pt";
            e1.AddText("Hello World");
            e1.SetText(s1);
        }
    }
}