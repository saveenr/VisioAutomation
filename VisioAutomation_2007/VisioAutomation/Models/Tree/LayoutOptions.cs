using VA=VisioAutomation;

namespace VisioAutomation.Models.Tree
{
    public class LayoutOptions
    {
        public ConnectorType ConnectorType { get; set; }
        public LayoutDirection Direction { get; set; }
        public VA.Drawing.Size DefaultNodeSize { get; set; }
        public VA.DOM.ShapeCells ConnectorCells { get; set; }
        
        public LayoutOptions()
        {
            DefaultNodeSize = new VA.Drawing.Size(2, 0.5);
            Direction = LayoutDirection.Down;
            this.ConnectorType = ConnectorType.DynamicConnector;
        }
    }
}