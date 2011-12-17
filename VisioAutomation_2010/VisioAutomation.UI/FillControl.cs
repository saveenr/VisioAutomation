using System.Windows.Forms;

namespace VisioAutomation.UI
{
    public partial class FillControl : UserControl
    {
        public BasicFillControl ShadowDef
        {
            get { return this.basicFillControlShadow; }
        }

        public BasicFillControl FillDef
        {
            get { return this.basicFillControlFill; }
        }

        public FillControl()
        {
            InitializeComponent();
        }

    }
}