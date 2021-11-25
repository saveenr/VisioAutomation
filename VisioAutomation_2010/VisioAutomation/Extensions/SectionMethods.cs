namespace VisioAutomation.Extensions;

public static class SectionMethods
{
    public static IEnumerable<IVisio.Row> ToEnumerable(this IVisio.Section section)
    {
        // Section object: http://msdn.microsoft.com/en-us/library/ms408988(v=office.12).aspx
        return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToEnumerable(() => section.Count, i => section[(short)i]);
    }

    public static List<IVisio.Row> ToList(this IVisio.Section section)
    {
        return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToList(() => section.Count, i => section[(short)i]);
    }
}