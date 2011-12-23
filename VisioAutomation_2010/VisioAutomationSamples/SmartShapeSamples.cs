using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

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

            string pinyf = string.Format("GUARD({0}!PinY)", bkname);
            string pinxf = string.Format("GUARD({0}!PinX-{0}!LocPinX+LocPinX)", bkname);
            string heightf = string.Format("GUARD({0}!Height)", bkname);
            string widthf = string.Format("GUARD({0}!Width*(PAGENUMBER()/PAGECOUNT()))", bkname);

            fg.CellsU["PinY"].Formula = pinyf;
            fg.CellsU["PinX"].Formula = pinxf;
            fg.CellsU["Height"].Formula = heightf;
            fg.CellsU["Width"].Formula = widthf;

            page.Application.ActiveWindow.SelectAll();
            var group = page.Application.ActiveWindow.Selection.Group();

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