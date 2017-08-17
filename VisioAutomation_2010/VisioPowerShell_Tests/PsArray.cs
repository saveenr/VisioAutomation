namespace VisioPowerShell_Tests
{
    public class PsArray
    {
        public static PsArray<T> From<T>(params T[] items)
        {
            return new PsArray<T>(items);
        }
    }

    public class PsArray<T>
    {
        private T[] Items;

        public PsArray()
        {
            this.Items = null;
        }

        public PsArray(T item)
        {
            this.Items = new T[]{ item };
        }

        public PsArray(params T[] items)
        {
            this.Items = items;
        }

        public T[] Array => this.Items;
    }
}