namespace VisioAutomation.Extensions;

public static class LayersMethods
{
    public static IEnumerable<IVisio.Layer> ToEnumerable(this IVisio.Layers layers)
    {
        return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToEnumerable(() => layers.Count, i => layers[i + 1]);
    }

    public static List<IVisio.Layer> ToList(this IVisio.Layers layers)
    {
        return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToList(() => layers.Count, i => layers[i + 1]);
    }
}