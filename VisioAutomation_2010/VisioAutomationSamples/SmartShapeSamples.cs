using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomationSamples
{
    public static class SmartShapeSamples
    {
        public static void ProgressBar()
        {
            var page_a = SampleEnvironment.Application.ActiveDocument.Pages.Add();


            // Draw some shapes
            var background = page_a.DrawRectangle(0, 0, 5, 1);
            var progress = page_a.DrawRectangle(0, 0, 1, 1);

            var background_fmt = new VA.Shapes.FormatCells();
            background_fmt.FillForegnd= "rgb(240,240,240)";
            background_fmt.LineColor = "rgb(100,100,100)";


            var progress_fmt = new VA.Shapes.FormatCells();
            progress_fmt.FillForegnd = "rgb(100,150,240)";
            progress_fmt.LineColor = "rgb(100,100,100)";

            // group the two shapes together
            page_a.Application.ActiveWindow.SelectAll();
            var group = page_a.Application.ActiveWindow.Selection.Group();

            // Set the progress shape update itself based on its position
            string bkname = background.NameID;
            var xform = new VA.Shapes.XFormCells();
            xform.PinX = string.Format("GUARD({0}!PinX-{0}!LocPinX+LocPinX)", bkname);
            xform.PinY = $"GUARD({bkname}!PinY)";
            xform.Width = $"GUARD({bkname}!Width*(PAGENUMBER()/PAGECOUNT()))";
            xform.Height = $"GUARD({bkname}!Height)"; 

            var update = new VA.ShapeSheet.Update();
            update.SetFormulas(progress.ID16, xform);
            update.SetFormulas(progress.ID16, background_fmt);
            update.SetFormulas(progress.ID16, progress_fmt);
            update.Execute(page_a);

            var markup1 = new VA.Text.Markup.TextElement();
            markup1.AddField(VA.Text.Markup.FieldConstants.PageName);
            markup1.AddText(" (");
            markup1.AddField(VA.Text.Markup.FieldConstants.PageNumber);
            markup1.AddText(" of ");
            markup1.AddField(VA.Text.Markup.FieldConstants.NumberOfPages);
            markup1.AddText(") ");
            markup1.SetText(group);
        }
    }
}