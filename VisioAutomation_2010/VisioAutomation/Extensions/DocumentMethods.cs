namespace VisioAutomation.Extensions;

public static class DocumentMethods
{
    public static void Close(this IVisio.Document doc, bool force_close)
    {
        Documents.DocumentHelper.Close(doc, force_close);
    }

    public static IEnumerable<IVisio.Document> ToEnumerable(this IVisio.Documents documents)
    {
        return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToEnumerable(() => documents.Count, i => documents[i + 1]); ;
    }

    public static List<IVisio.Document> ToList(this IVisio.Documents documents)
    {
        return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToList(() => documents.Count, i => documents[i + 1]); ;
    }

    public static IVisio.Document OpenStencil(this IVisio.Documents documents, string filename)
    {
        return Documents.DocumentHelper.OpenStencil(documents, filename);
    }
}