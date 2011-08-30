using IVisio= Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Experimental.SimpleTextDoc
{
    public class TextPage
    {
        public string Title;
        public string Body;
        public string Name;

        public IVisio.Page VisioPage;
        public IVisio.Shape VisioTitleShape;
        public IVisio.Shape VisioBodyShape;
    }
}