namespace VisioAutomation.ShapeSheet.Query;

public class Rows<T> : IEnumerable<Row<T>>
{

    private readonly List<Row<T>> _list;

    internal Rows(int capacity)
    {
        this._list = new List<Row<T>>(capacity);
    }

    public IEnumerator<Row<T>> GetEnumerator()
    {
        return this._list.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    internal void Add(Row<T> r)
    {
        this._list.Add(r);
    }

    internal void AddRange(IEnumerable<Row<T>> rows)
    {
        this._list.AddRange(rows);
    }

    public int Count => this._list.Count;

    public Row<T> this[int index] => this._list[index];
}