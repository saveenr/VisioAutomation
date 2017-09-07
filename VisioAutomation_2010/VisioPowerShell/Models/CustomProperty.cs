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


        internal CustomProperty(int id, string propname, CustomPropertyCells propcells)
        {
            this.ShapeID = id;
            this.Name = propname;
            this.Value = propcells.Value.Value;
            this.Format = propcells.Format.Value;
            this.Invisible = propcells.Invisible.Value;
            this.Label = propcells.Label.Value;
            this.LangId = propcells.LangID.Value;
            this.Prompt = propcells.Prompt.Value;
            this.SortKey = propcells.SortKey.Value;
            this.Type = propcells.Type.Value;
            this.Ask = propcells.Ask.Value;
            this.Calendar = propcells.Calendar.Value;
        }
    }
}