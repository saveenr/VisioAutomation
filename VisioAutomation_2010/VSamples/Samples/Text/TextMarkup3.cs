namespace VSamples.Samples.Text
{
    public  class TextMarkup3 : SampleMethodBase
    {
        public override void RunSample()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var segoe_ui = page.Document.Fonts["Segoe UI"];

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);
            var e1 = new VisioAutomation.Models.Text.Element();
            e1.AddText("When, from behind that craggy steep\n");
            e1.AddText("till then the horizon’s bound\n");

            var f2 = new TextFmt { HAlign = VisioScripting.Models.AlignmentHorizontal.Left, CharStyle = VisioAutomation.Models.Text.CharStyle.Italic };
            var f3 = new TextFmt { FontID = segoe_ui.ID, HAlign = VisioScripting.Models.AlignmentHorizontal.Center, CharStyle = VisioAutomation.Models.Text.CharStyle.Bold };
            var f4 = new TextFmt { HAlign = VisioScripting.Models.AlignmentHorizontal.Right, CharStyle = VisioAutomation.Models.Text.CharStyle.Italic };

            var e2 = e1.AddElementEx("a huge peak, black and huge\n", f2);
            var e3 = e1.AddElementEx("As if with voluntary power instinct\n", f3);
            var e4 = e1.AddElementEx("Upreared its head.\n", f4 );
            e1.AddText("-William Wordsworth, the Prelude");
            e1.SetText(s1);
        }
    }
}