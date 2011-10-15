using Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.PageLayout
{
    public class BasePageLayoutConfiguration
    {
        public double Spacing;
        public ConnectorStyle ConnectorStyle;
        public ConnectorAppearance ConnectorAppearance;
        public bool ResizePageToFit;
        public VA.Drawing.Size Border = new VA.Drawing.Size(0.5,0.5);
        public virtual void SetPageCells( VisioAutomation.Pages.PageCells pagecells)
        {
        }

        public void Apply(Page page)
        {
            var pagecells = new VA.Pages.PageCells();
            this.SetPageCells(pagecells);

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            pagecells.Apply(update);
            var pagesheet = page.PageSheet;
            update.Execute(pagesheet);
            page.Layout();

            if (this.ResizePageToFit)
            {
                if (this.Border.Height > 0 || this.Border.Width > 0)
                {
                    page.ResizeToFitContents(this.Border);
                }
                else
                {
                    page.ResizeToFitContents();                    
                }
            }
        }
    }
}