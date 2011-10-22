using VA=VisioAutomation;

namespace VisioAutomation.Layout.Tree
{
    public class LayoutOptions
    {
        public LayoutOptions()
        {
            DefaultNodeSize = new VA.Drawing.Size(2, 0.5);
            Direction = LayoutDirection.Down;
            this.ConnectorType = ConnectorType.DynamicConnector;
        }

        public ConnectorType ConnectorType { get; set; }      
        public LayoutDirection Direction { get; set; }
        public VA.Drawing.Size DefaultNodeSize { get; set; }

        public VA.DOM.ShapeCells ConnectorShapeCells { get; set; }
    }
}