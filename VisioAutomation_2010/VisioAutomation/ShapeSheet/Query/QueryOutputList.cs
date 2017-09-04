using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQueryOutputList<T> : List<QueryOutputCells<T>>
    {
        internal CellQueryOutputList() : base()
        {
        }
    }

    public class SectionQueryOutputList<T> : List<QueryOutputSections<T>>
    {
        internal SectionQueryOutputList() : base()
        {
        }
    }
}