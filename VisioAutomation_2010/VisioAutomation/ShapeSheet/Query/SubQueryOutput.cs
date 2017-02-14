using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SubQueryOutput<T>
    {
        public readonly List<SubQueryOutputRow<T>> Rows;

        internal SubQueryOutput(int capacity)
        {
            this.Rows = new List<SubQueryOutputRow<T>>(capacity);
        }
    }
}