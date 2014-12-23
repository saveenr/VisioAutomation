using VA=VisioAutomation;

namespace VisioPowerShell
{
    public class ConnectionPointValues
    {
        public readonly int ShapeID;
        public readonly string Type;
        public readonly string X;
        public readonly string Y;
        public readonly string DirX;
        public readonly string DirY;

        internal ConnectionPointValues(int shapeid, VA.Shapes.Connections.ConnectionPointCells point)
        {
            this.ShapeID = shapeid;
            this.Type = point.Type.Formula.Value;
            this.X = point.X.Formula.Value;
            this.Y = point.Y.Formula.Value;
            this.DirX = point.DirX.Formula.Value;
            this.DirY = point.DirY.Formula.Value;
        }
    }
}