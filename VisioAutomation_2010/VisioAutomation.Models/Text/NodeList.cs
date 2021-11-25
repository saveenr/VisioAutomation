using System.Collections.Generic;
using System.Collections;

namespace VisioAutomation.Models.Text;

public class NodeList<T> : IEnumerable<T> where T : Node
{
    private readonly Node _parent;
    private readonly List<T> _items;

    public NodeList(Node parentnode)
    {
        this._parent = parentnode;
        this._items = new List<T>(0);
    }

    public IEnumerator<T> GetEnumerator()
    {
        foreach (var i in this._items)
        {
            yield return i;
        }
    }

    IEnumerator IEnumerable.GetEnumerator()     
    {                                           
        return this.GetEnumerator();
    }

    public void Add(T item)
    {
        if (item.Parent != null)
        {
            if (item.Parent == this._parent)
            {
                throw new System.ArgumentException("already a child of parent");
            }
            else
            {
                throw new System.ArgumentException("already a child of another node");
            }
        }

        item.Parent = this._parent;
        this._items.Add(item);
    }

    public void Remove(T item)
    {
        if (item.Parent == null)
        {
            throw new System.ArgumentException("node does not have parent");
        }

        if (item.Parent != this._parent)
        {
            throw new System.ArgumentException("already a child of a different parent");
        }

        this._items.Remove(item);
    }
        
    public int Count
    {
        get { return this._items.Count; }
    }
        
    public T this[int i]
    {
        get
        {
            return this._items[i];
        }
    }
}