using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.LayoutStyles
{
    public abstract class LayoutStyleBase
    {
        public ConnectorStyle ConnectorStyle { get; set; }
        public ConnectorAppearance ConnectorAppearance { get; set; }
        public double AvenueSizeX { get; set; }
        public double AvenueSizeY { get; set; }

        protected LayoutStyleBase()
        {
            this.AvenueSizeX = 0.375;
            this.AvenueSizeY = 0.375;
        }

        protected virtual void SetPageCells(VisioAutomation.Pages.PageLayoutCells page_layout_cells)
        {
            page_layout_cells.AvenueSizeX = this.AvenueSizeX;
            page_layout_cells.AvenueSizeY = this.AvenueSizeY;
            page_layout_cells.LineRouteExt = (int) LayoutStyleBase._connector_appearance_to_line_route_ext(this.ConnectorAppearance);

            var rs = this.ConnectorsStyleToRouteStyle();
            if (rs.HasValue)
            {
                page_layout_cells.RouteStyle = (int) rs.Value;
            }
        }

        private static IVisio.VisCellVals _connector_appearance_to_line_route_ext(ConnectorAppearance con_appearance)
        {
            if (con_appearance == ConnectorAppearance.Default)
            {
                return IVisio.VisCellVals.visLORouteExtDefault;
            }
            else if (con_appearance == ConnectorAppearance.Straight)
            {
                return IVisio.VisCellVals.visLORouteExtStraight;
            }
            else if (con_appearance == ConnectorAppearance.Curved)
            {
                return IVisio.VisCellVals.visLORouteExtNURBS;
            }
            else
            {
                throw new System.ArgumentOutOfRangeException(nameof(con_appearance));
            }
        }

        protected virtual IVisio.VisCellVals? ConnectorsStyleToRouteStyle()
        {
            var cs = this.ConnectorStyle;
            if (cs == ConnectorStyle.RightAngle)
            {
                return IVisio.VisCellVals.visLORouteRightAngle;
            }
            else if (cs == ConnectorStyle.Straight)
            {
                return IVisio.VisCellVals.visLORouteStraight;
            }
            else if (cs == ConnectorStyle.CenterToCenter)
            {
                return IVisio.VisCellVals.visLORouteCenterToCenter;
            }
            else if (cs == ConnectorStyle.Network)
            {
                return IVisio.VisCellVals.visLORouteNetwork;
            }
            else
            {
                return null;
            }
        }

        protected IVisio.VisCellVals ConnectorsStyleAndDirectionToRouteStyle(ConnectorStyle con_style, LayoutDirection dir)
        {
            if (con_style == ConnectorStyle.Flowchart)
            {
                if (dir == LayoutDirection.BottomToTop)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartSN;
                }
                else if (dir == LayoutDirection.TopToBottom)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartNS;
                }
                else if (dir == LayoutDirection.LeftToRight)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartWE;
                }
                else if (dir == LayoutDirection.RightToLeft)
                {
                    return IVisio.VisCellVals.visLORouteFlowchartEW;
                }
            }
            else if (con_style == ConnectorStyle.OrganizationChart)
            {
                if (dir == LayoutDirection.BottomToTop)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartSN;
                }
                else if (dir == LayoutDirection.TopToBottom)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartNS;
                }
                else if (dir == LayoutDirection.LeftToRight)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartWE;
                }
                else if (dir == LayoutDirection.RightToLeft)
                {
                    return IVisio.VisCellVals.visLORouteOrgChartEW;
                }
            }
            else if (con_style == ConnectorStyle.Simple)
            {
                if (dir == LayoutDirection.BottomToTop)
                {
                    return IVisio.VisCellVals.visLORouteSimpleSN;
                }
                else if (dir == LayoutDirection.TopToBottom)
                {
                    return IVisio.VisCellVals.visLORouteSimpleNS;
                }
                else if (dir == LayoutDirection.LeftToRight)
                {
                    return IVisio.VisCellVals.visLORouteSimpleWE;
                }
                else if (dir == LayoutDirection.RightToLeft)
                {
                    return IVisio.VisCellVals.visLORouteSimpleEW;
                }
            }
            throw new System.ArgumentOutOfRangeException(nameof(con_style));
        }

        public void Apply(IVisio.Page page)
        {
            var page_layout_cells = new VisioAutomation.Pages.PageLayoutCells();
            this.SetPageCells(page_layout_cells);

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(page_layout_cells);

            writer.CommitFormulas(page.PageSheet);
            page.Layout();
        }
    }
}