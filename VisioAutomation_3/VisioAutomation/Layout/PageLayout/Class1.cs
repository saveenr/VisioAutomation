
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layout.PageLayout
{
    public enum PageLayoutStyle
    {
        Radial,
        Flowchart,
        Circular,
        CompactTree,
        Hierarchy
    }

    public enum FlowchartDirection
    {
        TopToBottom,
        BottomToTop,
        LeftToRight,
        RightToLeft
    }

    public enum HierarchyDirection
    {
        TopToBottom,
        BottomToTop,
        LeftToRight,
        RightToLeft
    }

    public enum CompactTreeDirection
    {
        DownThenRight,
        RightThenDown,
        RightThenUp,
        UpThenRigtht,
        UpThenLeft,
        LeftThenUp,
        LeftThenDown,
        DownThenLeft
    }

    public enum HierarchyHorizontalAlignment
    {
        Left,
        Center,
        Right
    }

    public enum HierarchyVerticalAlignment
    {
        Top,
        Middle,
        Bottom
    }


    public enum ConnectorStyle
    {
        RightAngle,
        Straight,
        CenterToCenter,
        Flowchart,
        Tree,
        OrganizationChart,
        Simple,
        SimpleHorizontalVertical,
        SimpleVerticalHorizontal
    }

    public enum ConnectorAppearance
    {
        Straight,
        Curved
    }

    public class BasePageLayoutConfiguration
    {
        public double Spacing;
        public ConnectorStyle ConnectorStyle;
        public ConnectorAppearance ConnectorAppearance;

        public virtual void SetPageCells( VisioAutomation.Pages.PageCells pagecells)
        {
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
        }
    }

    public class RadialConfiguration : BasePageLayoutConfiguration
    {
        public override void SetPageCells( VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceDefault;
        }        
    }

    public class FlowChartConfiguration: BasePageLayoutConfiguration
    {
        public FlowchartDirection Direction;

        public override void SetPageCells( VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            if (this.Direction == FlowchartDirection.TopToBottom)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceTopToBottom;                
            }
            else if (this.Direction == FlowchartDirection.LeftToRight)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceLeftToRight;
            }
            else if (this.Direction == FlowchartDirection.BottomToTop)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceBottomToTop;
            }
            else if (this.Direction == FlowchartDirection.RightToLeft)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceRightToLeft;
            }
            else
            {
                throw new VA.AutomationException();
            }
        }        

    }

    public class CircularConfiguration : BasePageLayoutConfiguration
    {
        public override void SetPageCells( VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceCircular;
        }        

    }

    public class CompactTreeConfiguration : BasePageLayoutConfiguration
    {
        public CompactTreeDirection Direction;

        public override void SetPageCells( VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            if (this.Direction == CompactTreeDirection.DownThenRight)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceCompactDownRight;
            }
            else if (this.Direction == CompactTreeDirection.RightThenDown)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceCompactRightDown;
            }
            else if (this.Direction == CompactTreeDirection.RightThenUp)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceCompactRightUp;
            }
            else if (this.Direction == CompactTreeDirection.UpThenRigtht)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceCompactUpRight;
            }
            else if (this.Direction == CompactTreeDirection.UpThenLeft)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceCompactUpLeft;
            }
            else if (this.Direction == CompactTreeDirection.LeftThenUp)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceCompactLeftUp;
            }
            else if (this.Direction == CompactTreeDirection.LeftThenDown)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceCompactLeftDown;
            }
            else if (this.Direction == CompactTreeDirection.DownThenLeft)
            {
                pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceCompactDownLeft;
            }
            else
            {
                throw new VA.AutomationException();
            }
        }        

    }

    public class HierarchyConfiguration : BasePageLayoutConfiguration
    {
        public HierarchyDirection Direction;
        public HierarchyHorizontalAlignment HorizontalAlignment;
        public HierarchyVerticalAlignment VerticalAlignment;

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            if (this.Direction == HierarchyDirection.BottomToTop)
            {
                if (this.HorizontalAlignment == HierarchyHorizontalAlignment.Left)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopLeft;                    
                }
                else if (this.HorizontalAlignment == HierarchyHorizontalAlignment.Center)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopCenter;
                }
                else if (this.HorizontalAlignment == HierarchyHorizontalAlignment.Right)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopRight;
                }
            }
            else if (this.Direction == HierarchyDirection.TopToBottom)
            {
                if (this.HorizontalAlignment== HierarchyHorizontalAlignment.Left)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomLeft;
                }
                else if (this.HorizontalAlignment == HierarchyHorizontalAlignment.Center)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomCenter;
                }
                else if (this.HorizontalAlignment == HierarchyHorizontalAlignment.Right)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomRight;
                }
            }
            else if (this.Direction == HierarchyDirection.LeftToRight)
            {
                if (this.VerticalAlignment == HierarchyVerticalAlignment.Top)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightTop;
                }
                else if (this.VerticalAlignment == HierarchyVerticalAlignment.Middle)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightMiddle;
                }
                else if (this.VerticalAlignment == HierarchyVerticalAlignment.Bottom)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightBottom;
                }
            }
            else if (this.Direction == HierarchyDirection.RightToLeft)
            {
                if (this.VerticalAlignment == HierarchyVerticalAlignment.Top)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyRightToLeftTop;
                }
                else if (this.VerticalAlignment == HierarchyVerticalAlignment.Middle)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyRightToLeftMiddle;
                }
                else if (this.VerticalAlignment == HierarchyVerticalAlignment.Bottom)
                {
                    pagecells.PlaceStyle = (int)IVisio.VisCellVals.visPLOPlaceHierarchyRightToLeftBottom;
                }
            }
            else
            {
                throw new VA.AutomationException();
            }

        }        

    }

}
