using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages.PageLayout
{
    public abstract class Layout
    {
        public LayoutStyle LayoutStyle { get; set; }
        public ConnectorStyle ConnectorStyle { get; set; }
        public ConnectorAppearance ConnectorAppearance { get; set; }
        public Drawing.Size AvenueSize { get; set; }

        protected Layout()
        {
            this.AvenueSize = new Drawing.Size(0.375, 0.375);
        }

        protected virtual void SetPageCells(PageLayoutFormulas pagecells)
        {
            pagecells.AvenueSizeX = this.AvenueSize.Width;
            pagecells.AvenueSizeY = this.AvenueSize.Height;
            pagecells.LineRouteExt = (int) Layout.ConnectorAppearanceToLineRouteExt(this.ConnectorAppearance);

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

        protected IVisio.VisCellVals ConnectorsStyleAndDirectionToRouteStyle(ConnectorStyle cs, Direction dir)
        {
            if (cs == ConnectorStyle.Flowchart)
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
            else if (cs == ConnectorStyle.OrganizationChart)
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
            else if (cs == ConnectorStyle.Simple)
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
            throw new System.ArgumentOutOfRangeException(nameof(cs));
        }

        public void Apply(IVisio.Page page)
        {
            var pagecells = new PageLayoutFormulas();
            this.SetPageCells(pagecells);

            var writer = new FormulaWriterSRC();
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.AvenueSizeX,pagecells.AvenueSizeX);
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.AvenueSizeY, pagecells.AvenueSizeY);
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.LineRouteExt, pagecells.LineRouteExt);
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.RouteStyle, pagecells.RouteStyle);
            writer.SetFormula(VisioAutomation.ShapeSheet.SRCConstants.PlaceStyle, pagecells.PlaceStyle);

            var pagesheet = page.PageSheet;

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(pagesheet);
            writer.Commit(surface);
            page.Layout();
        }
    }
}