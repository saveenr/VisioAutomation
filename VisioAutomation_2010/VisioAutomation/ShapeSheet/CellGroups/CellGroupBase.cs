using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupBase
    {
        public abstract IEnumerable<SrcFormulaPair> SrcFormulaPairs { get; }
    }
}