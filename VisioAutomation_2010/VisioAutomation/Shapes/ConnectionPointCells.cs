using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public class ConnectionPointCells : CellGroup
    {
        public CellValueLiteral X { get; set; }
        public CellValueLiteral Y { get; set; }
        public CellValueLiteral DirX { get; set; }
        public CellValueLiteral DirY { get; set; }
        public CellValueLiteral Type { get; set; }

        public override IEnumerable<NamedSrcValuePair> NamedSrcValuePairs
        {
            get
            {
                yield return NamedSrcValuePair.Create(nameof(this.X), SrcConstants.ConnectionPointX, this.X);
                yield return NamedSrcValuePair.Create(nameof(this.Y), SrcConstants.ConnectionPointY, this.Y);
                yield return NamedSrcValuePair.Create(nameof(this.DirX), SrcConstants.ConnectionPointDirX, this.DirX);
                yield return NamedSrcValuePair.Create(nameof(this.DirY), SrcConstants.ConnectionPointDirY, this.DirY);
                yield return NamedSrcValuePair.Create(nameof(this.Type), SrcConstants.ConnectionPointType, this.Type);
            }
        }


    }
}