using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioPowerShell
{
    public class CustomPropertyValues
    {
        public readonly int ShapeID;
        public readonly string Name;
        public readonly string Value;
        public readonly string Format;
        public readonly string Invisible;
        public readonly string Label;
        public readonly string LangId;
        public readonly string Prompt;
        public readonly string SortKey;
        public readonly string Type;
        public readonly string Verify;
        public readonly string Calendar;

        internal CustomPropertyValues(int id, string propname, VA.Shapes.CustomProperties.CustomPropertyCells propcells)
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
            this.Calendar = propcells.Calendar.Formula.Value;
        }
    }
}