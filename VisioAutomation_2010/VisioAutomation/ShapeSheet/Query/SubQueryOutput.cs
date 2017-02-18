
namespace VisioAutomation.ShapeSheet.Query
{
    public class SubQueryOutput<T>
    {
        public readonly SubQueryOutputRowCollection<T> Rows;

        internal SubQueryOutput(int capacity)
        {
            this.Rows = new SubQueryOutputRowCollection<T>(capacity);
        }
    }
}