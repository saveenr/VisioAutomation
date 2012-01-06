using VisioAutomation.Text.Markup;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

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

            var e1 = new VA.Text.Markup.TextElement();
            e1.CharacterFormat.Color = new VA.Drawing.ColorRGB(0xff0000);
            e1.CharacterFormat.Font = tnr.Name;
            e1.CharacterFormat.FontSize = 20;
            e1.AppendText("Hello World");
            e1.SetText(s1);
        }

        public static void TextMarkup12()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);
            var tnr = page.Document.Fonts["Times New Roman"];

            var e1 = new VA.Text.Markup.TextElement();
            e1.CharacterFormat.Color = new VA.Drawing.ColorRGB(0xff0000);
            e1.CharacterFormat.Font = tnr.Name;
            e1.CharacterFormat.FontSize = 20;
            e1.AppendText("Hello ");

            var e2 = e1.AddElementEx("World", null, null, null, null, VA.Text.CharStyle.Italic); 
            e1.SetText(s1);
        }

        public static void TextMarkup13()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);
            var e1 = new VA.Text.Markup.TextElement();
            e1.AppendText("When, from behind that craggy steep\n");
            e1.AppendText("till then the horizon’s bound\n");
            var e2 = e1.AddElementEx("a huge peak, black and huge\n", null, null, null, VA.Drawing.AlignmentHorizontal.Left, VA.Text.CharStyle.Italic);
            var e3 = e1.AddElementEx("As if with voluntary power instinct\n", "Segoe UI", null, null, VA.Drawing.AlignmentHorizontal.Center, VA.Text.CharStyle.Bold);
            var e4 = e1.AddElementEx("Upreared its head.\n", null, null, null, VA.Drawing.AlignmentHorizontal.Right, VA.Text.CharStyle.Italic);
            e1.AppendText("-William Wordsworth, the Prelude");
            e1.SetText(s1);
        }

        public static void TextMarkup14()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);
            var e1 = new VA.Text.Markup.TextElement();
            e1.AppendText("This shape is ");
            e1.AppendField(VA.Text.Markup.FieldConstants.Width);
            e1.AppendText("inches wide by ");
            e1.AppendField(VA.Text.Markup.FieldConstants.Height);
            e1.AppendText("inches tall.");
            e1.SetText(s1);
        }
    }
}