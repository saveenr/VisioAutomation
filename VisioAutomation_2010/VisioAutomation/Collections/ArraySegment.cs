namespace VisioAutomation.Collections;

public struct ArraySegment<T> : IEnumerable<T>
{
    private readonly T[] _array;
    private readonly int _offset;
    private readonly int _count;

    public ArraySegment(T[] array, int offset, int count)
    {
        this._array = array;
        this._offset = offset;
        this._count = count;
    }

    public T this[int index]
    {
        get { return get_value_at_index(index); }
        set { set_value_at_index(index, value); }
    }

    private void set_value_at_index(int index, T value)
    {
        validate_index(index);
        this._array[this._offset + index] = value;
    }

    private T get_value_at_index(int index)
    {
        validate_index(index);

        var value = this._array[this._offset + index];
        return value;
    }

    private void validate_index(int index)
    {
        if ((index < 0) && (index >= this._count))
        {
            throw new System.ArgumentOutOfRangeException(nameof(index));
        }
    }

    public IEnumerator<T> GetEnumerator()
    {
        for (int i = 0; i < this._count; i++)
        {
            yield return this._array[_offset + i];
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public int Count => this._count;

    public int Offset => this._offset;
}