using VisioAutomation.Shapes;

namespace VisioPowerShell.Models
{
    public class CustomProperty
    {
        public readonly int ShapeID;
        public readonly string Name;

        public readonly string Ask;
        public readonly string Calendar;
        public readonly string Format;
        public readonly string Invisible;
        public readonly string Label;
        public readonly string LangId;
        public readonly string Prompt;
        public readonly string SortKey;
        public readonly string Type;
        public readonly string Value;


        internal CustomProperty(int id, string name, CustomPropertyCells cells)
        {
            this.ShapeID = id;
            this.Name = name;
            this.Value = cells.Value.Value;
            this.Format = cells.Format.Value;
            this.Invisible = cells.Invisible.Value;
            this.Label = cells.Label.Value;
            this.LangId = cells.LangID.Value;
            this.Prompt = cells.Prompt.Value;
            this.SortKey = cells.SortKey.Value;
            this.Type = cells.Type.Value;
            this.Ask = cells.Ask.Value;
            this.Calendar = cells.Calendar.Value;
        }
    }
}