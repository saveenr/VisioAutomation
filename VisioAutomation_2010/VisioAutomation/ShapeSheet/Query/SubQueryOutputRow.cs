using System.Collections;


namespace VisioAutomation.ShapeSheet.Query
{
    public struct SubQueryOutputRow<T>  
    {
        public readonly CellRange<T> Cells;

        internal SubQueryOutputRow(CellRange<T> cells)
        {
            this.Cells = cells;
        }
    }
}