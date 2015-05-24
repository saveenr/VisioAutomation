using VACUSTPROP=VisioAutomation.Shapes.CustomProperties;

namespace VisioPowerShell.Model
{
    public class CustomPropertyValues
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


        internal CustomPropertyValues(int id, string propname, VACUSTPROP.CustomPropertyCells propcells)
        {
            this.ShapeID = id;
            this.Name = propname;
            this.Value = propcells.Value.Formula.Value;
            this.Format = propcells.Format.Formula.Value;
            this.Invisible = propcells.Invisible.Formula.Value;
            this.Label = propcells.Label.Formula.Value;
            this.LangId = propcells.LangId.Formula.Value;
            this.Prompt = propcells.Prompt.Formula.Value;
            this.SortKey = propcells.SortKey.Formula.Value;
            this.Type = propcells.Type.Formula.Value;
            this.Ask = propcells.Ask.Formula.Value;
            this.Calendar = propcells.Calendar.Formula.Value;
        }
    }
}