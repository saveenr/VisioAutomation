using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layout.PageLayout
{
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
            pagecells.PlaceStyle = (int) GetPlaceStyle(this.Direction);                
        }        

        private static IVisio.VisCellVals GetPlaceStyle(FlowchartDirection dir)
        {
            if (dir == FlowchartDirection.TopToBottom)
            {
                return IVisio.VisCellVals.visPLOPlaceTopToBottom;
            }
            else if (dir == FlowchartDirection.LeftToRight)
            {
                return IVisio.VisCellVals.visPLOPlaceLeftToRight;
            }
            else if (dir == FlowchartDirection.BottomToTop)
            {
                return IVisio.VisCellVals.visPLOPlaceBottomToTop;
            }
            else if (dir == FlowchartDirection.RightToLeft)
            {
                return IVisio.VisCellVals.visPLOPlaceRightToLeft;
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
        public HierarchyDirection Direction;
        public HierarchyHorizontalAlignment HorizontalAlignment;
        public HierarchyVerticalAlignment VerticalAlignment;

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) GetPlaceStyle(this.Direction, this.HorizontalAlignment, this.VerticalAlignment);                    

        }

        private static IVisio.VisCellVals GetPlaceStyle(HierarchyDirection dir, HierarchyHorizontalAlignment halign , HierarchyVerticalAlignment valign)
        {
            if (dir == HierarchyDirection.BottomToTop)
            {
                if (halign == HierarchyHorizontalAlignment.Left)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopLeft;
                }
                else if (halign == HierarchyHorizontalAlignment.Center)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopCenter;
                }
                else if (halign == HierarchyHorizontalAlignment.Right)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyBottomToTopRight;
                }
            }
            else if (dir == HierarchyDirection.TopToBottom)
            {
                if (halign == HierarchyHorizontalAlignment.Left)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomLeft;
                }
                else if (halign == HierarchyHorizontalAlignment.Center)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomCenter;
                }
                else if (halign == HierarchyHorizontalAlignment.Right)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyTopToBottomRight;
                }
            }
            else if (dir == HierarchyDirection.LeftToRight)
            {
                if (valign == HierarchyVerticalAlignment.Top)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightTop;
                }
                else if (valign == HierarchyVerticalAlignment.Middle)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightMiddle;
                }
                else if (valign == HierarchyVerticalAlignment.Bottom)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyLeftToRightBottom;
                }
            }
            else if (dir == HierarchyDirection.RightToLeft)
            {
                if (valign == HierarchyVerticalAlignment.Top)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyRightToLeftTop;
                }
                else if (valign == HierarchyVerticalAlignment.Middle)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyRightToLeftMiddle;
                }
                else if (valign == HierarchyVerticalAlignment.Bottom)
                {
                    return  IVisio.VisCellVals.visPLOPlaceHierarchyRightToLeftBottom;
                }
                else
                {
                    throw new VA.AutomationException();
                }
            }
            throw new VA.AutomationException();

        }


    }

}
