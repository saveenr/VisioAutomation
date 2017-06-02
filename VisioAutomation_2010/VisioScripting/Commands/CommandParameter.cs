namespace VisioScripting.Commands
{
    public class CommandParameter
    {
        public readonly string Name;
        public readonly System.Type Type;
        public readonly string TypeDisplayName;

        internal CommandParameter(string name, System.Type type, string typename)
        {
            this.Name = name;
            this.Type = type;
            this.TypeDisplayName = typename;
        }
    }
}