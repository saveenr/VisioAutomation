using VAM = VisioAutomation.Models;
using VSM = VisioScripting.Models;

namespace VSamples.Text
{
    public  class TextMarkpSamples1 : SampleMethodBase
    {
        public override void RunSample()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);
            var tnr = page.Document.Fonts["Times New Roman"];

            var e1 = new VAM.Text.Element();
            var color_red = new VAM.Color.ColorRgb(0xff0000);
            e1.CharacterFormatting.Color = color_red.ToFormula();
            e1.CharacterFormatting.Font = tnr.ID;
            e1.CharacterFormatting.Font = "20pt";
            e1.AddText("Hello World");
            e1.SetText(s1);
        }
    }
    public  class TextMarkpSamples2 : SampleMethodBase
    {

        public override void RunSample()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);
            var tnr = page.Document.Fonts["Times New Roman"];

            var e1 = new VAM.Text.Element();
            var color_red = new VAM.Color.ColorRgb(0xff0000);
            e1.CharacterFormatting.Color = color_red.ToFormula();
            e1.CharacterFormatting.Font = tnr.ID;
            e1.CharacterFormatting.Font = "20pt";
            e1.AddText("Hello ");

            var f2 = new TextFmt { CharStyle = VAM.Text.CharStyle.Italic };
            var e2 = e1.AddElementEx("World", f2 );
            e1.SetText(s1);
        }
    }
    public  class TextMarkpSamples3 : SampleMethodBase
    {
        public override void RunSample()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var segoe_ui = page.Document.Fonts["Segoe UI"];

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);
            var e1 = new VAM.Text.Element();
            e1.AddText("When, from behind that craggy steep\n");
            e1.AddText("till then the horizon’s bound\n");

            var f2 = new TextFmt { HAlign = VSM.AlignmentHorizontal.Left, CharStyle = VAM.Text.CharStyle.Italic };
            var f3 = new TextFmt { FontID = segoe_ui.ID, HAlign = VSM.AlignmentHorizontal.Center, CharStyle = VAM.Text.CharStyle.Bold };
            var f4 = new TextFmt { HAlign = VSM.AlignmentHorizontal.Right, CharStyle = VAM.Text.CharStyle.Italic };

            var e2 = e1.AddElementEx("a huge peak, black and huge\n", f2);
            var e3 = e1.AddElementEx("As if with voluntary power instinct\n", f3);
            var e4 = e1.AddElementEx("Upreared its head.\n", f4 );
            e1.AddText("-William Wordsworth, the Prelude");
            e1.SetText(s1);
        }
    }
    public  class TextMarkpSamples4 : SampleMethodBase
    {
        public override void RunSample()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);
            var e1 = new VAM.Text.Element();
            e1.AddText("This shape is ");
            e1.AddField(VAM.Text.FieldConstants.Width);
            e1.AddText("inches wide by ");
            e1.AddField(VAM.Text.FieldConstants.Height);
            e1.AddText("inches tall.");
            e1.SetText(s1);
        }
    }
    public  class TextMarkpSamples5 : SampleMethodBase
    {
        public override void RunSample()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8.5, 11);

            var e1 = new VAM.Text.Element();
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