using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.Outputs
{
    public class SubQueryOutput<T>
    {
        public SubQuery Column { get; internal set; }
        public readonly List<SubQueryOutputRow<T>> Rows;

        internal SubQueryOutput(int capacity)
        {
            this.Rows = new List<SubQueryOutputRow<T>>(capacity);
        }
    }
}