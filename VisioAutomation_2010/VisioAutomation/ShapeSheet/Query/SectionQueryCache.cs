
namespace VisioAutomation.ShapeSheet.Query;

internal class SectionQueryCache
{
    private readonly List<ShapeCache> _list;

    public SectionQueryCache()
    {
        this._list = new List<ShapeCache>();
    }

    public SectionQueryCache(int capacity)
    {
        this._list = new List<ShapeCache>(capacity);
    }

    public void Add(ShapeCache item)
    {
        this._list.Add(item);
    }

    public int Count
    {
        get
        {
            return this._list.Count;
        }
    }

    public IEnumerable<ShapeCache> ShapeCacheItems
    {
        get
        {
            return this._list;
        }
    }

    public ShapeCache this[int index]
    {
        get
        {
            return this._list[index];
        }
    }

    public int CountCells()
    {
        // Count the cells not in sections
        int count = 0;
        foreach (var section_info in this.ShapeCacheItems)
        {
            count += section_info.CountCells();
        }

        return count;
    }
}