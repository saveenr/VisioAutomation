using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.PageLayouts
{
    public abstract class LayoutBase
    {
        public LayoutStyle LayoutStyle { get; set; }
        public ConnectorStyle ConnectorStyle { get; set; }
        public ConnectorAppearance ConnectorAppearance { get; set; }
        public Geometry.Size AvenueSize { get; set; }

        protected LayoutBase()
        {
            this.AvenueSize = new Geometry.Size(0.375, 0.375);
        }

        protected virtual void SetPageCells(VisioAutomation.Pages.PageLayoutCells page_layout_cells)
        {
            page_layout_cells.AvenueSizeX = this.AvenueSize.Width;
            page_layout_cells.AvenueSizeY = this.AvenueSize.Height;
            page_layout_cells.LineRouteExt = (int) LayoutBase.ConnectorAppearanceToLineRouteExt(this.ConnectorAppearance);

            var rs = this.ConnectorsStyleToRouteStyle();
            if (rs.HasValue)
            {
                page_layout_cells.RouteStyle = (int) rs.Value;
            }
        }

        private static IVisio.VisCellVals ConnectorAppearanceToLineRouteExt(ConnectorAppearance con_appearance)
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
            page_layout_cells.SetFormulas(writer);

            writer.Commit(page.PageSheet);
            page.Layout();
        }
    }
}