using System.Collections;


namespace VisioAutomation.ShapeSheet.Query
{
    public struct SubQueryOutputRow<T>  
    {
        public readonly VisioAutomation.Utilities.ArraySegment<T> Cells;

        internal SubQueryOutputRow(VisioAutomation.Utilities.ArraySegment<T> cells)
        {
            this.Cells = cells;
        }
    }
}