using Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Layout.PageLayout
{
    public class BasePageLayoutConfiguration
    {
        public double Spacing;
        public ConnectorStyle ConnectorStyle;
        public ConnectorAppearance ConnectorAppearance;

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
        }
    }
}