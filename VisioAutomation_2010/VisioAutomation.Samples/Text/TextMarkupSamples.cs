using VA = VisioAutomation;

namespace VisioAutomationSamples
{
    public static class TextMarkpSamples
    {
        public static void TextMarkup11()
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

        public static void TextMarkup12()
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
            e1.AddText("Hello ");

            var e2 = e1.AddElementEx("World", null, null, null, null, VA.Models.Text.CharStyle.Italic); 
            e1.SetText(s1);
        }

        public static void TextMarkup13()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var segoe_ui = page.Document.Fonts["Segoe UI"];

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);
            var e1 = new VisioAutomation.Models.Text.Element();
            e1.AddText("When, from behind that craggy steep\n");
            e1.AddText("till then the horizon’s bound\n");
            var e2 = e1.AddElementEx("a huge peak, black and huge\n", null, null, null, VisioScripting.Models.AlignmentHorizontal.Left, VA.Models.Text.CharStyle.Italic);
            var e3 = e1.AddElementEx("As if with voluntary power instinct\n", segoe_ui.ID, null, null, VisioScripting.Models.AlignmentHorizontal.Center, VA.Models.Text.CharStyle.Bold);
            var e4 = e1.AddElementEx("Upreared its head.\n", null, null, null, VisioScripting.Models.AlignmentHorizontal.Right, VA.Models.Text.CharStyle.Italic);
            e1.AddText("-William Wordsworth, the Prelude");
            e1.SetText(s1);
        }

        public static void TextMarkup14()
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

        public static void TextMarkup5()
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