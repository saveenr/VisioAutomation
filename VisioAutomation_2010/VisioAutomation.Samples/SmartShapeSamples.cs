using VA = VisioAutomation;

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

            var background_fmt = new VA.Shapes.ShapeFormatCells();
            background_fmt.FillForegnd= "rgb(240,240,240)";
            background_fmt.LineColor = "rgb(100,100,100)";


            var progress_fmt = new VA.Shapes.ShapeFormatCells();
            progress_fmt.FillForegnd = "rgb(100,150,240)";
            progress_fmt.LineColor = "rgb(100,100,100)";

            // group the two shapes together
            page_a.Application.ActiveWindow.SelectAll();
            var group = page_a.Application.ActiveWindow.Selection.Group();

            // Set the progress shape update itself based on its position
            string bkname = background.NameID;
            var xform = new VA.Shapes.XFormCells();
            xform.PinX = string.Format("GUARD({0}!PinX-{0}!LocPinX+LocPinX)", bkname);
            xform.PinY = string.Format("GUARD({0}!PinY)", bkname);
            xform.Width = string.Format("GUARD({0}!Width*(PAGENUMBER()/PAGECOUNT()))", bkname);
            xform.Height = string.Format("GUARD({0}!Height)", bkname); 

            var writer = new VisioAutomation.ShapeSheet.ShapeSheetWriter();
            xform.SetFormulas(progress.ID16, writer);
            background_fmt.SetFormulas(progress.ID16, writer);
            progress_fmt.SetFormulas(progress.ID16, writer);

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(page_a);
            writer.Commit(surface);

            var markup1 = new VisioAutomation.Models.Text.TextElement();
            markup1.AddField(VisioAutomation.Models.Text.FieldConstants.PageName);
            markup1.AddText(" (");
            markup1.AddField(VisioAutomation.Models.Text.FieldConstants.PageNumber);
            markup1.AddText(" of ");
            markup1.AddField(VisioAutomation.Models.Text.FieldConstants.NumberOfPages);
            markup1.AddText(") ");
            markup1.SetText(group);
        }
    }
}