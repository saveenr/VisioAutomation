using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeLayout
{
    public class RadialLayoutOptions : LayoutOptions
    {
        public RadialLayoutOptions() :
            base()
        {
            this.LayoutStyle = LayoutStyle.Radial;
            this.ConnectorStyle = ConnectorStyle.RightAngle;
        }

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceDefault;
        }
    }

    public class FlowChartLayoutOptions : LayoutOptions
    {
        public Direction Direction;

        public FlowChartLayoutOptions() :
            base()
        {
            this.LayoutStyle = LayoutStyle.Flowchart;
            this.ConnectorStyle = ConnectorStyle.Flowchart;
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

    public class CircularLayoutOptions : LayoutOptions
    {
        public CircularLayoutOptions() :
            base()
        {
            this.LayoutStyle = LayoutStyle.Circular;
            this.ConnectorStyle = ConnectorStyle.CenterToCenter;

        }

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceCircular;
        }


    }

    public class CompactTreeOptions : LayoutOptions
    {
        public CompactTreeDirection Direction;

        public CompactTreeOptions() :
            base()
        {
            this.LayoutStyle = LayoutStyle.CompactTree;
            this.ConnectorStyle = ConnectorStyle.OrganizationChart;

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

    public class HierarchyLayoutOptions : LayoutOptions
    {
        public Direction Direction;
        public HorizontalAlignment HorizontalAlignment;
        public VerticalAlignment VerticalAlignment;

        public HierarchyLayoutOptions() :
            base()
        {
            this.LayoutStyle = LayoutStyle.Hierarchy;
            this.ConnectorStyle = ConnectorStyle.OrganizationChart;
        }

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) GetPlaceStyle(this.Direction, this.HorizontalAlignment, this.VerticalAlignment);
        }


        private static IVisio.VisCellVals GetPlaceStyle(Direction dir, HorizontalAlignment halign, VerticalAlignment valign)
        {
            if (dir == Direction.BottomToTop)
            {
                if (halign == HorizontalAlignment.Left)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopLeft;
                }
                else if (halign == HorizontalAlignment.Center)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopCenter;
                }
                else if (halign == HorizontalAlignment.Right)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopRight;
                }
            }
            else if (dir == Direction.TopToBottom)
            {
                if (halign == HorizontalAlignment.Left)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomLeft;
                }
                else if (halign == HorizontalAlignment.Center)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomCenter;
                }
                else if (halign == HorizontalAlignment.Right)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomRight;
                }
            }
            else if (dir == Direction.LeftToRight)
            {
                if (valign == VerticalAlignment.Top)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightTop;
                }
                else if (valign == VerticalAlignment.Middle)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightMiddle;
                }
                else if (valign == VerticalAlignment.Bottom)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightBottom;
                }
            }
            else if (dir == Direction.RightToLeft)
            {
                if (valign == VerticalAlignment.Top)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyRightToLeftTop;
                }
                else if (valign == VerticalAlignment.Middle)
                {
                    return IVisio.VisCellVals.visPLOPlaceHierarchyRightToLeftMiddle;
                }
                else if (valign == VerticalAlignment.Bottom)
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