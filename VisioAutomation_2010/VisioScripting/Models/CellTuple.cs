namespace VisioScripting.Models
{
    public struct CellTuple
    {
        public string Name;
        public VisioAutomation.ShapeSheet.Src Src;
        public string Formula;

        public CellTuple(string name, VisioAutomation.ShapeSheet.Src src, string formula)
        {
            this.Name = name;
            this.Src = src;
            this.Formula = formula;
        }
    }
}