using VisioAutomation.ShapeSheet.Writers;
using VA = VisioAutomation;

namespace VSamples
{
    public  class ProgressBarX : SampleMethodBase
    {
        public override void RunSample()
        {
            var page_a = SampleEnvironment.Application.ActiveDocument.Pages.Add();


            // Draw some shapes
            var background = page_a.DrawRectangle(0, 0, 5, 1);
            var progress = page_a.DrawRectangle(0, 0, 1, 1);

            var background_fmt = new VA.Shapes.ShapeFormatCells();
            background_fmt.FillForeground= "rgb(240,240,240)";
            background_fmt.LineColor = "rgb(100,100,100)";


            var progress_fmt = new VA.Shapes.ShapeFormatCells();
            progress_fmt.FillForeground = "rgb(100,150,240)";
            progress_fmt.LineColor = "rgb(100,100,100)";

            // group the two shapes together
            page_a.Application.ActiveWindow.SelectAll();
            var group = page_a.Application.ActiveWindow.Selection.Group();

            // Set the progress shape update itself based on its position
            string bkname = background.NameID;
            var xfrm = new VA.Shapes.ShapeXFormCells();
            xfrm.PinX = string.Format("GUARD({0}!PinX-{0}!LocPinX+LocPinX)", bkname);
            xfrm.PinY = string.Format("GUARD({0}!PinY)", bkname);
            xfrm.Width = string.Format("GUARD({0}!Width*(PAGENUMBER()/PAGECOUNT()))", bkname);
            xfrm.Height = string.Format("GUARD({0}!Height)", bkname); 

            var writer = new SidSrcWriter();

            writer.SetValues(progress.ID16, xfrm);
            writer.SetValues(progress.ID16, background_fmt);
            writer.SetValues(progress.ID16, progress_fmt);

            writer.Commit(page_a, VA.Core.CellValueType.Formula);

            var markup1 = new VisioAutomation.Models.Text.Element();
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