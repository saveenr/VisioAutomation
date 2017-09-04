using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQueryOutputList<T> : List<CellQueryOutput<T>>
    {
        internal CellQueryOutputList() : base()
        {
        }
    }

    public class SectionQueryOutputList<T> : List<SectionsQueryOutput<T>>
    {
        internal SectionQueryOutputList() : base()
        {
        }
    }
}