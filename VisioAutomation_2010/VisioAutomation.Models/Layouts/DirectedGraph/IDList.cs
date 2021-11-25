using System.Collections;

namespace VisioAutomation.Models.Layouts.DirectedGraph;

public class IDList<T> : IEnumerable<T> where T : class
{
    private readonly Dictionary<string, T> _items;

    public IDList()
    {
        this._items = new Dictionary<string, T>();
    }

    public void Add(string id, T item)
    {
        this._items.Add(id, item);
    }

    public T this[string index]
    {
        get { return this._items[index]; }
    }

    public int Count
    {
        get { return this._items.Count; }
    }

    public IEnumerator<T> GetEnumerator()
    {
        foreach (var i in this._items.Values)
        {
            yield return i;
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public bool ContainsID(string id)
    {
        return this._items.ContainsKey(id);
    }

    public IEnumerable<string> IDs
    {
        get
        {
            foreach (var id in this._items.Keys)
            {
                yield return id;
            }
        }
    }

    public T Find(string id)
    {
        T item = null;
        if (this._items.TryGetValue(id, out item))
        {
            return item;
        }

        return null;
    }
}