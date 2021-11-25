namespace VisioScripting.Commands;

public class CommandSet
{
    // Keep a reference back to the parent client. This gives access to all other commands
    // for a the current context
    protected readonly Client _client;

    internal CommandSet(Client client)
    {
        this._client = client;
    }

    internal static IEnumerable<Command> GetCommands(System.Type mytype)
    {
        var cmdsettype = typeof(CommandSet);

        if (!cmdsettype.IsAssignableFrom(mytype))
        {
            string msg = string.Format("{0} must derive from {1}", mytype.Name, cmdsettype.Name);
            throw new System.ArgumentException(msg,nameof(mytype));
        }

        var methods = mytype.GetMethods().Where(m => m.IsPublic && !m.IsStatic);

        foreach (var method in methods)
        {
            // Skip some method names
            switch (method.Name)
            {
                case "ToString":
                case "GetHashCode":
                case "GetType":
                case "Equals":
                    continue;
            }

            var cmd = new Command(method);
            yield return cmd;
        }
    }
}