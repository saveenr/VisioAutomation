using VisioAutomation.Shapes;

namespace VisioPowerShell.Models
{
    public class ConnectionPoint
    {
        public readonly int ShapeID;
        public readonly string Type;
        public readonly string X;
        public readonly string Y;
        public readonly string DirX;
        public readonly string DirY;

        internal ConnectionPoint(int shapeid, ConnectionPointCells point)
        {
            this.ShapeID = shapeid;
            this.Type = point.Type.Value;
            this.X = point.X.Value;
            this.Y = point.Y.Value;
            this.DirX = point.DirX.Value;
            this.DirY = point.DirY.Value;
        }
    }
}