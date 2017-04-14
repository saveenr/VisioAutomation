namespace VisioPowerShell.Models
{
    public class UserDefinedCell
    {
        public readonly int ShapeID;
        public readonly string Name;
        public readonly string Value;
        public readonly string Prompt;

        public UserDefinedCell(int id, string name, string value, string prompt)
        {
            this.ShapeID = id;
            this.Name = name;
            this.Value = value;
            this.Prompt = prompt;
        }
    }
}