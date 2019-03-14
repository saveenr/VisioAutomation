using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class ConnectionPointCells : CellGroup
    {
        public CellValueLiteral X { get; set; }
        public CellValueLiteral Y { get; set; }
        public CellValueLiteral DirX { get; set; }
        public CellValueLiteral DirY { get; set; }
        public CellValueLiteral Type { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.ConnectionPointX, this.X);
                yield return SrcValuePair.Create(SrcConstants.ConnectionPointY, this.Y);
                yield return SrcValuePair.Create(SrcConstants.ConnectionPointDirX, this.DirX);
                yield return SrcValuePair.Create(SrcConstants.ConnectionPointDirY, this.DirY);
                yield return SrcValuePair.Create(SrcConstants.ConnectionPointType, this.Type);
            }
        }


    }
}