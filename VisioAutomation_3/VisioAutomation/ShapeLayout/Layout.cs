using Microsoft.Office.Interop.Visio;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.ShapeLayout
{
    public abstract class Layout
    {
        public LayoutStyle LayoutStyle;
        public ConnectorStyle ConnectorStyle;
        public ConnectorAppearance ConnectorAppearance;
        public bool ResizePageToFit;
        public VA.Drawing.Size Border = new VA.Drawing.Size(0.5,0.5);
        public VA.Drawing.Size AvenueSize = new VA.Drawing.Size(0.375, 0.375);

        protected Layout()
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
            if (ca == VA.ShapeLayout.ConnectorAppearance.Default)
            {
                return IVisio.VisCellVals.visLORouteExtDefault;
            }
            else if (ca == VA.ShapeLayout.ConnectorAppearance.Straight)
            {
                return IVisio.VisCellVals.visLORouteExtStraight;
            }
            else if (ca == VA.ShapeLayout.ConnectorAppearance.Curved)
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
            if (cs == VA.ShapeLayout.ConnectorStyle.RightAngle)
            {
                return IVisio.VisCellVals.visLORouteRightAngle;
            }
            else if (cs == VA.ShapeLayout.ConnectorStyle.Straight)
            {
                return IVisio.VisCellVals.visLORouteStraight;
            }
            else if (cs == VA.ShapeLayout.ConnectorStyle.CenterToCenter)
            {
                return IVisio.VisCellVals.visLORouteCenterToCenter;
            }
            else if (cs == VA.ShapeLayout.ConnectorStyle.Network)
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
            if (cs == VA.ShapeLayout.ConnectorStyle.Flowchart)
            {
                if (dir  == Direction.BottomToTop)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartSN;
                }
                else if (dir  == Direction.TopToBottom)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartNS;
                }
                else if (dir  == Direction.LeftToRight)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartWE;
                }
                else if (dir  == Direction.RightToLeft)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartEW;
                }
            }
            else if (cs == VA.ShapeLayout.ConnectorStyle.OrganizationChart)
            {
                if (dir  == Direction.BottomToTop)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartSN;
                }
                else if (dir  == Direction.TopToBottom)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartNS;
                }
                else if (dir  == Direction.LeftToRight)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartWE;
                }
                else if (dir  == Direction.RightToLeft)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartEW;
                }

            }
            else if (cs == VA.ShapeLayout.ConnectorStyle.Simple)
            {
                if (dir  == Direction.BottomToTop)
                {
                    return IVisio.VisCellVals.visLORouteSimpleSN;
                }
                else if (dir  == Direction.TopToBottom)
                {
                    return IVisio.VisCellVals.visLORouteSimpleNS;
                }
                else if (dir  == Direction.LeftToRight)
                {
                    return IVisio.VisCellVals.visLORouteSimpleWE;
                }
                else if (dir  == Direction.RightToLeft)
                {
                    return IVisio.VisCellVals.visLORouteSimpleEW;
                }
            }
            throw new VA.AutomationException();
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