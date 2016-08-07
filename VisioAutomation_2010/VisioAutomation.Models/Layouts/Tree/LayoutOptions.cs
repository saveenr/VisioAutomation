using VA=VisioAutomation;

namespace VisioAutomation.Models.Layouts.Tree
{
    public class LayoutOptions
    {
        public ConnectorType ConnectorType { get; set; }
        public LayoutDirection Direction { get; set; }
        public VA.Drawing.Size DefaultNodeSize { get; set; }
        public DOM.ShapeCells ConnectorCells { get; set; }
        
        public LayoutOptions()
        {
            this.DefaultNodeSize = new VA.Drawing.Size(2, 0.5);
            this.Direction = LayoutDirection.Down;
            this.ConnectorType = ConnectorType.DynamicConnector;
        }
    }
}