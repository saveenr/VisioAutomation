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

            var rs = this.ConnectorsStyleToRouteStyle();
            if (rs.HasValue)
            {
                pagecells.RouteStyle = (int)rs.Value;                
            }
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

        protected virtual IVisio.VisCellVals? ConnectorsStyleToRouteStyle()
        {
            var cs = this.ConnectorStyle;
            if (cs == VA.Layout.PageLayout.ConnectorStyle.RightAngle)
            {
                return IVisio.VisCellVals.visLORouteRightAngle;
            }
            else if (cs == VA.Layout.PageLayout.ConnectorStyle.Straight)
            {
                return IVisio.VisCellVals.visLORouteStraight;
            }
            else if (cs == VA.Layout.PageLayout.ConnectorStyle.CenterToCenter)
            {
                return IVisio.VisCellVals.visLORouteCenterToCenter;
            }
            else if (cs == VA.Layout.PageLayout.ConnectorStyle.Network)
            {
                return IVisio.VisCellVals.visLORouteNetwork;
            }
            else
            {
                return null;
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