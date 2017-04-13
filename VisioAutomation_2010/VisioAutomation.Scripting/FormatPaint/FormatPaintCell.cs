namespace VisioAutomation.Scripting.FormatPaint
{
    public class FormatPaintCell
    {
        public FormatPaintCategory PaintCategory { get; }
        public VisioAutomation.ShapeSheet.Src Src { get; private set; }
        public string Name;

        public string Result { get; set; }
        public string Formula { get; set; }

        public FormatPaintCell(VisioAutomation.ShapeSheet.Src src, string name, FormatPaintCategory paint_category)
        {
            this.PaintCategory = paint_category;
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

        public bool MatchesCategory(FormatPaintCategory paint_category)
        {
            return ((this.PaintCategory & paint_category) != 0);
        }
    }
}