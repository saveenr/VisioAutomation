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
            this.Value = propcells.Value.ValueF;
            this.Format = propcells.Format.ValueF;
            this.Invisible = propcells.Invisible.ValueF;
            this.Label = propcells.Label.ValueF;
            this.LangId = propcells.LangID.ValueF;
            this.Prompt = propcells.Prompt.ValueF;
            this.SortKey = propcells.SortKey.ValueF;
            this.Type = propcells.Type.ValueF;
            this.Ask = propcells.Ask.ValueF;
            this.Calendar = propcells.Calendar.ValueF;
        }
    }
}