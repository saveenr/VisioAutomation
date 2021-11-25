

namespace VisioPowerShell.Internal;

public struct CellTuple
{
    public string Name;
    public VASS.Src Src;
    public string Formula;

    public CellTuple(string name, VASS.Src src, string formula)
    {
        this.Name = name;
        this.Src = src;
        this.Formula = formula;
    }
}