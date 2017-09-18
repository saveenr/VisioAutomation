using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupBase
    {
        public abstract IEnumerable<SrcValuePair> SrcValuePairs { get; }
    }
}