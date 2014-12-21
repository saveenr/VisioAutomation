namespace VisioPowerShell
{
    public class UserDefinedCellvalues
    {
        public readonly int ShapeID;
        public readonly string Name;
        public readonly string Value;
        public readonly string Prompt;

        public UserDefinedCellvalues(int id, string name, string value, string prompt)
        {
            this.ShapeID = id;
            this.Name = name;
            this.Value = value;
            this.Prompt = prompt;
        }
    }
}