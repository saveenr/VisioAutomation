namespace VSamples.Samples.Text
{
    public  class TextMarkup5 : SampleMethodBase
    {
        public override void Run()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);

            var e1 = new VisioAutomation.Models.Text.Element();
            e1.ParagraphFormatting.HorizontalAlign = 0;
            var e2 = e1.AddElement("Hello Worldline1\nline2\nline3\n");
            e2.ParagraphFormatting.IndentFirst = "0.5pt";
            e2.ParagraphFormatting.IndentLeft= "0.25pt";
            var e3 = e1.AddElement("Goodbye\nline1\nline2\nline3");
            e3.ParagraphFormatting.IndentFirst = "1.0pt";
            e3.ParagraphFormatting.IndentLeft= "0.75pt";

            e1.SetText(s1);
        }
    }
}