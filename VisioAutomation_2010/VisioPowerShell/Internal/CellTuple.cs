namespace VisioPowerShell.Internal
{
    public struct CellTuple
    {
        public string Name;
        public VisioAutomation.Core.Src Src;
        public string Formula;

        public CellTuple(string name, VisioAutomation.Core.Src src, string formula)
        {
            this.Name = name;
            this.Src = src;
            this.Formula = formula;
        }
    }
}