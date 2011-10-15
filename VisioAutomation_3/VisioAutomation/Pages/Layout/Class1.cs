using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layout.PageLayout
{
    public class RadialConfiguration : BasePageLayoutConfiguration
    {
        public RadialConfiguration() :
            base()
        {
            this.LayoutStyle = VA.Layout.PageLayout.PageLayoutStyle.Radial;
            this.ConnectorStyle = VA.Layout.PageLayout.ConnectorStyle.RightAngle;
        }

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceDefault;
        }
    }

    public class FlowChartConfiguration : BasePageLayoutConfiguration
    {
        public Direction Direction;

        public FlowChartConfiguration() :
            base()
        {
            this.LayoutStyle = VA.Layout.PageLayout.PageLayoutStyle.Flowchart;
            this.ConnectorStyle = VA.Layout.PageLayout.ConnectorStyle.Flowchart;
        }

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) GetPlaceStyle(this.Direction);
        }

        private static IVisio.VisCellVals GetPlaceStyle(Direction dir)
        {
            if (dir == Direction.TopToBottom)
            {
                return IVisio.VisCellVals.visPLOPlaceTopToBottom;
            }
            else if (dir == Direction.LeftToRight)
            {
                return IVisio.VisCellVals.visPLOPlaceLeftToRight;
            }
            else if (dir == Direction.BottomToTop)
            {
                return IVisio.VisCellVals.visPLOPlaceBottomToTop;
            }
            else if (dir == Direction.RightToLeft)
            {
                return IVisio.VisCellVals.visPLOPlaceRightToLeft;
            }
            else
            {
                throw new VA.AutomationException();
            }
        }

        protected override IVisio.VisCellVals? ConnectorsStyleToRouteStyle()
        {
            var rs = base.ConnectorsStyleToRouteStyle();
            if (rs.HasValue)
            {
                return rs;
            }
            return this.ConnectorsStyleAndDirectionToRouteStyle(this.ConnectorStyle, this.Direction);
        }


    }

    public class CircularConfiguration : BasePageLayoutConfiguration
    {
        public CircularConfiguration() :
            base()
        {
            this.LayoutStyle = VA.Layout.PageLayout.PageLayoutStyle.Circular;
            this.ConnectorStyle = VA.Layout.PageLayout.ConnectorStyle.CenterToCenter;

        }

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceCircular;
        }


    }

    public class CompactTreeConfiguration : BasePageLayoutConfiguration
    {
        public CompactTreeDirection Direction;

        public CompactTreeConfiguration() :
            base()
        {
            this.LayoutStyle = VA.Layout.PageLayout.PageLayoutStyle.CompactTree;
            this.ConnectorStyle = VA.Layout.PageLayout.ConnectorStyle.OrganizationChart;

        }

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) GetPlaceStyle(this.Direction);
        }

        private static IVisio.VisCellVals GetPlaceStyle(CompactTreeDirection dir)
        {
            if (dir == CompactTreeDirection.DownThenRight)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactDownRight;
            }
            else if (dir == CompactTreeDirection.RightThenDown)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactRightDown;
            }
            else if (dir == CompactTreeDirection.RightThenUp)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactRightUp;
            }
            else if (dir == CompactTreeDirection.UpThenRigtht)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactUpRight;
            }
            else if (dir == CompactTreeDirection.UpThenLeft)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactUpLeft;
            }
            else if (dir == CompactTreeDirection.LeftThenUp)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactLeftUp;
            }
            else if (dir == CompactTreeDirection.LeftThenDown)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactLeftDown;
            }
            else if (dir == CompactTreeDirection.DownThenLeft)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactDownLeft;
            }
            else
            {
                throw new VA.AutomationException();
            }
        }


    }

    public class HierarchyConfiguration : BasePageLayoutConfiguration
    {
        public Direction Direction;
        public HierarchyHorizontalAlignment HorizontalAlignment;
        public HierarchyVerticalAlignment VerticalAlignment;

        public HierarchyConfiguration() :
            base()
        {
            this.LayoutStyle = VA.Layout.PageLayout.PageLayoutStyle.Hierarchy;
            this.ConnectorStyle = VA.Layout.PageLayout.ConnectorStyle.OrganizationChart;
        }

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) GetPlaceStyle(this.Direction, this.HorizontalAlignment, this.VerticalAlignment);
        }


        private static IVisio.VisCellVals GetPlaceStyle(Direction dir, HierarchyHorizontalAlignment halign, HierarchyVerticalAlignment valign)
        {
            if (dir == Direction.BottomToTop)
            {
                if (halign == HierarchyHorizontalAlignment.Left)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopLeft;
                }
                else if (halign == HierarchyHorizontalAlignment.Center)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopCenter;
                }
                else if (halign == HierarchyHorizontalAlignment.Right)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopRight;
                }
            }
            else if (dir == Direction.TopToBottom)
            {
                if (halign == HierarchyHorizontalAlignment.Left)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomLeft;
                }
                else if (halign == HierarchyHorizontalAlignment.Center)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomCenter;
                }
                else if (halign == HierarchyHorizontalAlignment.Right)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomRight;
                }
            }
            else if (dir == Direction.LeftToRight)
            {
                if (valign == HierarchyVerticalAlignment.Top)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightTop;
                }
                else if (valign == HierarchyVerticalAlignment.Middle)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightMiddle;
                }
                else if (valign == HierarchyVerticalAlignment.Bottom)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightBottom;
                }
            }
            else if (dir == Direction.RightToLeft)
            {
                if (valign == HierarchyVerticalAlignment.Top)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyRightToLeftTop;
                }
                else if (valign == HierarchyVerticalAlignment.Middle)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyRightToLeftMiddle;
                }
                else if (valign == HierarchyVerticalAlignment.Bottom)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyRightToLeftBottom;
                }
                else
                {
                    throw new VA.AutomationException();
                }
            }
            throw new VA.AutomationException();
        }

        protected override IVisio.VisCellVals? ConnectorsStyleToRouteStyle()
        {
            var rs = base.ConnectorsStyleToRouteStyle();
            if (rs.HasValue)
            {
                return rs;
            }
            return this.ConnectorsStyleAndDirectionToRouteStyle(this.ConnectorStyle, this.Direction);
        }
    }
}