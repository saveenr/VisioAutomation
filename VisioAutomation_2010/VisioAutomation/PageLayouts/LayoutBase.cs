using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.PageLayouts
{
    public abstract class LayoutBase
    {
        public LayoutStyle LayoutStyle { get; set; }
        public ConnectorStyle ConnectorStyle { get; set; }
        public ConnectorAppearance ConnectorAppearance { get; set; }
        public Drawing.Size AvenueSize { get; set; }

        protected LayoutBase()
        {
            this.AvenueSize = new Drawing.Size(0.375, 0.375);
        }

        protected virtual void SetPageCells(PageLayoutFormulas pagecells)
        {
            pagecells.AvenueSizeX = this.AvenueSize.Width;
            pagecells.AvenueSizeY = this.AvenueSize.Height;
            pagecells.LineRouteExt = (int) LayoutBase.ConnectorAppearanceToLineRouteExt(this.ConnectorAppearance);

            var rs = this.ConnectorsStyleToRouteStyle();
            if (rs.HasValue)
            {
                pagecells.RouteStyle = (int) rs.Value;
            }
        }

        private static IVisio.VisCellVals ConnectorAppearanceToLineRouteExt(ConnectorAppearance ca)
        {
            if (ca == ConnectorAppearance.Default)
            {
                return IVisio.VisCellVals.visLORouteExtDefault;
            }
            else if (ca == ConnectorAppearance.Straight)
            {
                return IVisio.VisCellVals.visLORouteExtStraight;
            }
            else if (ca == ConnectorAppearance.Curved)
            {
                return IVisio.VisCellVals.visLORouteExtNURBS;
            }
            else
            {
                throw new System.ArgumentOutOfRangeException(nameof(ca));
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

        protected IVisio.VisCellVals ConnectorsStyleAndDirectionToRouteStyle(ConnectorStyle cs, LayoutDirection dir)
        {
            if (cs == ConnectorStyle.Flowchart)
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
            else if (cs == ConnectorStyle.OrganizationChart)
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
            else if (cs == ConnectorStyle.Simple)
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
            throw new System.ArgumentOutOfRangeException(nameof(cs));
        }

        public void Apply(IVisio.Page page)
        {
            var pagecells = new PageLayoutFormulas();
            this.SetPageCells(pagecells);

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutAvenueSizeX,pagecells.AvenueSizeX);
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutAvenueSizeY, pagecells.AvenueSizeY);
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineRouteExt, pagecells.LineRouteExt);
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutRouteStyle, pagecells.RouteStyle);
            writer.SetFormula(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutPlaceStyle, pagecells.PlaceStyle);

            writer.Commit(page.PageSheet);
            page.Layout();
        }
    }
}