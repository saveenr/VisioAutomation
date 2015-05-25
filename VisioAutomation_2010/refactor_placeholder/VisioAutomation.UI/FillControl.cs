using System.Windows.Forms;

namespace VisioAutomation.UI
{
    public partial class FillControl : UserControl
    {
        public BasicFillControl ShadowDef { get; private set; }

        public BasicFillControl FillDef { get; private set; }

        public FillControl()
        {
            this.InitializeComponent();
        }

    }
}