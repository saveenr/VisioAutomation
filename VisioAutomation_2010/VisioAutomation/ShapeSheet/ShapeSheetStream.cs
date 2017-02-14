namespace VisioAutomation.ShapeSheet
{
    public class ShapeSheetStream
    {
        internal short[] short_array;

        internal ShapeSheetStream(short[] a)
        {
            this.short_array = a;
        }

        public bool IsEmpty => this.short_array.Length == 0;
    }
}