using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.DOM
{
    public class PageSettings
    {
        public Drawing.Size? Size { get; set; }
        public PageCells PageCells { get; set; }

        public PageSettings()
        {
            this.PageCells = new PageCells();
        }

        public PageSettings(VA.Drawing.Size size) :
            this()
        {
            this.Size = size;
        }

        public PageSettings(double w, double h) :
            this(new VA.Drawing.Size(w, h))
        {
        }

        public void Apply(IVisio.Page page)
        {
            if (this.Size.HasValue)
            {
                page.SetSize(this.Size.Value);
            }

            var page_sheet = page.PageSheet;
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            this.PageCells.Apply(update, (short)page_sheet.ID);
            update.Execute(page);
        }
    }
}