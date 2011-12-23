using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomationSamples
{
    public static class SmartShapeSamples
    {
        public static void ProgressBar()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var bk = page.DrawRectangle(0, 0, 5, 1);
            var fg = page.DrawRectangle(0, 0, 1, 1);

            string bkname = bk.NameID;

            var xform  = new VA.Layout.XFormCells();
            xform.PinX =  string.Format("GUARD({0}!PinY)", bkname);
            xform.PinY = string.Format("GUARD({0}!PinX-{0}!LocPinX+LocPinX)", bkname);
            xform.Width = string.Format("GUARD({0}!Height)", bkname);
            xform.Height = string.Format("GUARD({0}!Width*(PAGENUMBER()/PAGECOUNT()))", bkname);
            
            page.Application.ActiveWindow.SelectAll();
            var group = page.Application.ActiveWindow.Selection.Group();

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            xform.Apply(update);
            update.Execute(group);
            
            var markup1 = new VA.Text.Markup.TextElement();
            markup1.AppendField(VA.Text.Markup.Fields.PageName);
            markup1.AppendText(" (");
            markup1.AppendField(VA.Text.Markup.Fields.PageNumber);
            markup1.AppendText(" of ");
            markup1.AppendField(VA.Text.Markup.Fields.NumberOfPages);
            markup1.AppendText(") ");
            markup1.SetText(group);
        }
    }
}