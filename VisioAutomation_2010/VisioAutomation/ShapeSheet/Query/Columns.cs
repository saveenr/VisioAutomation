
namespace VisioAutomation.ShapeSheet.Query;

public class Columns : IEnumerable<Column> 
{
    protected IList<Column> _items;
    protected Dictionary<string, Column> _map_name_to_item;
    protected Dictionary<ShapeSheet.Src, Column> _dic_src_to_col;

    internal Columns() 
    {
        this._items = new List<Column>();
        this._map_name_to_item = new Dictionary<string, Column>();

    }

    public IEnumerator<Column> GetEnumerator()
    {
        return (this._items).GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public Column this[int index] => this._items[index];

    public Column this[string name] => this._map_name_to_item[name];

    public bool Contains(string name) => this._map_name_to_item.ContainsKey(name);

    protected string normalize_name(string name)
    {
        if (string.IsNullOrEmpty(name))
        {
            name = string.Format("Col{0}", this._items.Count);
        }
        return name;
    }

    public int Count => this._items.Count;

    protected void check_duplicate_column_name(string name)
    {
        if (this._map_name_to_item.ContainsKey(name))
        {
            throw new System.ArgumentException("Duplicate Column Name");
        }
    }

    protected void check_deplicate_src(Src src)
    {
        if (this._dic_src_to_col == null)
        {
            this._dic_src_to_col = new Dictionary<ShapeSheet.Src, Column>();
        }

        if (this._dic_src_to_col.ContainsKey(src))
        {
            string msg = string.Format("Duplicate {0}({1},{2},{3})", nameof(Src), src.Section, src.Row, src.Cell);
            throw new System.ArgumentException(msg);
        }
    }


    public Column Add(ShapeSheet.Src src, string name)
    {
        if (name == null)
        {
            throw new System.ArgumentNullException(nameof(name));
        }

        check_deplicate_src(src);
        string norm_name = this.normalize_name(name);
        check_duplicate_column_name(norm_name);

        int ordinal = this._items.Count;
        var col = new Column(ordinal, norm_name, src);
        this._items.Add(col);

        this._map_name_to_item[norm_name] = col;
        this._dic_src_to_col.Add(src, col);
        return col;
    }

}