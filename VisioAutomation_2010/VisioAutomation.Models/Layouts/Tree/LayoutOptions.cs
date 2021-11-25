
namespace VisioAutomation.Models.Layouts.Tree
{
    public class LayoutOptions
    {
        public ConnectorType ConnectorType { get; set; }
        public LayoutDirection Direction { get; set; }
        public VA.Geometry.Size DefaultNodeSize { get; set; }
        public Dom.ShapeCells ConnectorCells { get; set; }
        
        public LayoutOptions()
        {
            this.DefaultNodeSize = new VA.Geometry.Size(2, 0.5);
            this.Direction = LayoutDirection.Down;
            this.ConnectorType = ConnectorType.DynamicConnector;
        }
    }
}