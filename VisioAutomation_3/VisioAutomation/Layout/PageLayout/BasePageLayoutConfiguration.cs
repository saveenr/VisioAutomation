using Microsoft.Office.Interop.Visio;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.PageLayout
{
    public class BasePageLayoutConfiguration
    {
        public VA.Layout.PageLayout.PageLayoutStyle LayoutStyle;
        public ConnectorStyle ConnectorStyle;
        public ConnectorAppearance ConnectorAppearance;
        public bool ResizePageToFit;
        public VA.Drawing.Size Border = new VA.Drawing.Size(0.5,0.5);
        public VA.Drawing.Size AvenueSize = new VA.Drawing.Size(0.375, 0.375);

        protected BasePageLayoutConfiguration()
        {
            
        }

        public virtual void SetPageCells( VisioAutomation.Pages.PageCells pagecells)
        {
            pagecells.AvenueSizeX = this.AvenueSize.Width;
            pagecells.AvenueSizeY = this.AvenueSize.Height;
            pagecells.LineRouteExt = (int)ConnectorAppearanceToLineRouteExt(this.ConnectorAppearance);
        }

        private static IVisio.VisCellVals ConnectorAppearanceToLineRouteExt( ConnectorAppearance ca)
        {
            if (ca == VA.Layout.PageLayout.ConnectorAppearance.Default)
            {
                return IVisio.VisCellVals.visLORouteExtDefault;
            }
            else if (ca == VA.Layout.PageLayout.ConnectorAppearance.Straight)
            {
                return IVisio.VisCellVals.visLORouteExtStraight;
            }
            else if (ca == VA.Layout.PageLayout.ConnectorAppearance.Curved)
            {
                return IVisio.VisCellVals.visLORouteExtNURBS;
            }
            else
            {
                throw new VA.AutomationException();
            }
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