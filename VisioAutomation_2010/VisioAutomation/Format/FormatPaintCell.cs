using VA=VisioAutomation;

namespace VisioAutomation.Format
{
    public class FormatPaintCell
    {
        public VA.Format.FormatCategory Category { get; private set; }
        public VA.ShapeSheet.SRC SRC { get; private set; }
        public string Name;

        public string Result { get; set; }
        public string Formula { get; set; }

        public FormatPaintCell(VA.ShapeSheet.SRC src, string name, FormatCategory category)
        {
            this.Category = category;
            this.Name = name;
            this.SRC = src;
            this.Formula = null;
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