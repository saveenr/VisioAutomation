using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.PageLayout
{
    public abstract class Layout
    {
        public LayoutStyle LayoutStyle { get; set; }
        public ConnectorStyle ConnectorStyle { get; set; }
        public ConnectorAppearance ConnectorAppearance { get; set; }
        public bool ResizePageToFit { get; set; }
        public VA.Drawing.Size Border { get; set; }
        public VA.Drawing.Size AvenueSize { get; set; }

        protected Layout()
        {
            this.Border = new VA.Drawing.Size(0.5, 0.5);
            this.AvenueSize = new VA.Drawing.Size(0.375, 0.375);
        }

        public virtual void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            pagecells.AvenueSizeX = this.AvenueSize.Width;
            pagecells.AvenueSizeY = this.AvenueSize.Height;
            pagecells.LineRouteExt = (int) ConnectorAppearanceToLineRouteExt(this.ConnectorAppearance);

            var rs = this.ConnectorsStyleToRouteStyle();
            if (rs.HasValue)
            {
                pagecells.RouteStyle = (int) rs.Value;
            }
        }

        private static IVisio.VisCellVals ConnectorAppearanceToLineRouteExt(ConnectorAppearance ca)
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

        protected IVisio.VisCellVals ConnectorsStyleAndDirectionToRouteStyle(ConnectorStyle cs, Direction dir)
        {
            if (cs == VA.Layout.PageLayout.ConnectorStyle.Flowchart)
            {
                if (dir == Direction.BottomToTop)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartSN;
                }
                else if (dir == Direction.TopToBottom)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartNS;
                }
                else if (dir == Direction.LeftToRight)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartWE;
                }
                else if (dir == Direction.RightToLeft)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartEW;
                }
            }
            else if (cs == VA.Layout.PageLayout.ConnectorStyle.OrganizationChart)
            {
                if (dir == Direction.BottomToTop)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartSN;
                }
                else if (dir == Direction.TopToBottom)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartNS;
                }
                else if (dir == Direction.LeftToRight)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartWE;
                }
                else if (dir == Direction.RightToLeft)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartEW;
                }
            }
            else if (cs == VA.Layout.PageLayout.ConnectorStyle.Simple)
            {
                if (dir == Direction.BottomToTop)
                {
                    return IVisio.VisCellVals.visLORouteSimpleSN;
                }
                else if (dir == Direction.TopToBottom)
                {
                    return IVisio.VisCellVals.visLORouteSimpleNS;
                }
                else if (dir == Direction.LeftToRight)
                {
                    return IVisio.VisCellVals.visLORouteSimpleWE;
                }
                else if (dir == Direction.RightToLeft)
                {
                    return IVisio.VisCellVals.visLORouteSimpleEW;
                }
            }
            throw new VA.AutomationException();
        }

        public void Apply(IVisio.Page page)
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