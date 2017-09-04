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
            this.Type = point.Type.Formula;
            this.X = point.X.Formula;
            this.Y = point.Y.Formula;
            this.DirX = point.DirX.Formula;
            this.DirY = point.DirY.Formula;
        }
    }
}