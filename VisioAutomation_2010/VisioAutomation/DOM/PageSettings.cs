using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.DOM
{
    public class PageSettings
    {
        public PageCells PageCells { get; set; }

        public PageSettings()
        {
            this.PageCells = new PageCells();
        }

        public void Apply(IVisio.Page page)
        {
            var page_sheet = page.PageSheet;
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            this.PageCells.Apply(update, (short)page_sheet.ID);
            update.Execute(page);
        }
    }
}