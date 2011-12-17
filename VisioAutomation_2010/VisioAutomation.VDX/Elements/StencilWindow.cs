namespace VisioAutomation.VDX.Elements
{
    public class StencilWindow : Window
    {
        public int ParentWindowID { get; set; }
        public string Document { get; set; }
        public int StencilGroup { get; set; }
        public int StencilGroupPos { get; set; }

        public StencilWindow() :
            base()
        {
        }
    }
}