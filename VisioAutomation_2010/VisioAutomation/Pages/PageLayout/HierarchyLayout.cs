using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages.PageLayout
{
    public class HierarchyLayout : Layout
    {
        public LayoutDirection LayoutDirection { get; set; }
        public HorizontalAlignment HorizontalAlignment { get; set; }
        public VerticalAlignment VerticalAlignment { get; set; }

        public HierarchyLayout()
        {
            this.LayoutStyle = LayoutStyle.Hierarchy;
            this.ConnectorStyle = ConnectorStyle.OrganizationChart;
            this.HorizontalAlignment = HorizontalAlignment.Center;
            this.VerticalAlignment = VerticalAlignment.Middle;
        }

        protected override void SetPageCells(PageLayoutFormulas pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) HierarchyLayout.GetPlaceStyle(this.LayoutDirection, this.HorizontalAlignment, this.VerticalAlignment);
        }

        private static IVisio.VisCellVals GetPlaceStyle(LayoutDirection dir, HorizontalAlignment halign, VerticalAlignment valign)
        {
            if (dir == LayoutDirection.BottomToTop)
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
            else if (dir == LayoutDirection.TopToBottom)
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
            else if (dir == LayoutDirection.LeftToRight)
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
            else if (dir == LayoutDirection.RightToLeft)
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
                    throw new System.ArgumentOutOfRangeException(nameof(dir));
                }
            }
            throw new System.ArgumentOutOfRangeException(nameof(dir));
        }

        protected override IVisio.VisCellVals? ConnectorsStyleToRouteStyle()
        {
            var rs = base.ConnectorsStyleToRouteStyle();
            if (rs.HasValue)
            {
                return rs;
            }
            return this.ConnectorsStyleAndDirectionToRouteStyle(this.ConnectorStyle, this.LayoutDirection);
        }
    }
}