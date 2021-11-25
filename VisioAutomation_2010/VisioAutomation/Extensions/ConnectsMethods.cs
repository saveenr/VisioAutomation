namespace VisioAutomation.Extensions
{
    public static class ConnectsMethods
    {
        public static IEnumerable<IVisio.Connect> ToEnumerable(this IVisio.Connects connects)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToEnumerable(() => connects.Count, i => connects[i + 1]);
        }

        public static List<IVisio.Connect> ToList(this IVisio.Connects connects)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToList(() => connects.Count, i => connects[i + 1]);
        }
    }
}
