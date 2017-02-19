namespace VisioAutomation.Utilities
{
    public class ArraySegmentBuilder<T>
    {
        private T[] array;
        private int pos;
        public int Count;

        public ArraySegmentBuilder(T[] array)
        {
            this.array = array;
            this.pos = 0;
            this.Count = 0;
        }

        public VisioAutomation.Utilities.ArraySegment<T> GetNextSegment(int size)
        {
            var seg = new VisioAutomation.Utilities.ArraySegment<T>(this.array, this.pos, size);
            this.pos += size;
            this.Count += size;
            return seg;
        }
    }
}