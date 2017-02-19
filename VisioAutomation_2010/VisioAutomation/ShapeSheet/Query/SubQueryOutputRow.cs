using System.Collections;


namespace VisioAutomation.ShapeSheet.Query
{
    public struct SubQueryOutputRow<T>  
    {
        public readonly T[] Cells;

        internal SubQueryOutputRow(T[] cells)
        {
            if (cells == null)
            {
                throw new System.ArgumentNullException(nameof(cells));
            }

            this.Cells = cells;
        }
    }
}