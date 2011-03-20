using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Layout.MSAGL
{
    public class Shape : Node
    {
        public Shape(string id)
        {
            this.ID = id;
        }

        public string StencilName { get; set; }
        public string MasterName { get; set; }
        public string URL { get; set; }
        public VA.Drawing.Size? Size { get; set; }
    }
}