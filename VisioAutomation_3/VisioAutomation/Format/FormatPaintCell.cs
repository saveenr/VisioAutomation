using VA=VisioAutomation;

namespace VisioAutomation.Format
{
    public class FormatPaintCell
    {
        public VA.Format.FormatCategory Category { get; private set; }
        public VA.ShapeSheet.SRC SRC { get; private set; }
        public double? Result { get; set; }
        public string Formula { get; set; }

        public FormatPaintCell(VA.ShapeSheet.SRC src, FormatCategory category)
        {
            this.Category = category;
            this.SRC = src;
            this.Result = null;
        }

        public void Clear()
        {
            this.Result = null;
            this.Formula = null;
        }

        public bool MatchesCategory(VA.Format.FormatCategory category)
        {
            return ((this.Category & category) != 0);
        }
    }
}