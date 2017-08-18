namespace VisioPowerShell_Tests.Framework
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
        private readonly T[] _array;

        public PsArray()
        {
            this._array = null;
        }

        public PsArray(params T[] items)
        {
            this._array = items;
        }

        public T[] Array => this._array;
    }
}