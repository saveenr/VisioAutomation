namespace VisioAutomation.Scripting.FormatPaint
{
    public class FormatPaintCell
    {
        public FormatCategory Category { get; }
        public VisioAutomation.ShapeSheet.Src Src { get; private set; }
        public string Name;

        public string Result { get; set; }
        public string Formula { get; set; }

        public FormatPaintCell(VisioAutomation.ShapeSheet.Src src, string name, FormatCategory category)
        {
            this.Category = category;
            this.Name = name;
            this.Src = src;
            this.Formula = null;
            this.Result = null;
        }

        public void Clear()
        {
            this.Result = null;
            this.Formula = null;
        }

        public bool MatchesCategory(FormatCategory category)
        {
            return ((this.Category & category) != 0);
        }
    }
}