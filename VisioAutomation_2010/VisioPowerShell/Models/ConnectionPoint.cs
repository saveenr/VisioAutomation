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
            this.Type = point.Type.ValueF;
            this.X = point.X.ValueF;
            this.Y = point.Y.ValueF;
            this.DirX = point.DirX.ValueF;
            this.DirY = point.DirY.ValueF;
        }
    }
}