using System.Windows.Forms;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioPowerTools
{
    public partial class FormStackingTool : Form
    {
        public FormStackingTool()
        {
            InitializeComponent();
        }

        private void FormArrangeTool_Load(object sender, System.EventArgs e)
        {
            this.comboBoxSnapDelta.Text = Globals.VisioPowerToolsAddIn.addinprefs.SnapUnit.ToString();
        }

        public void buttonLayoutE2EH_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Arrange.Stack(VA.Drawing.Axis.XAxis, Globals.VisioPowerToolsAddIn.addinprefs.SnapUnit);
        }

        public void buttonLayoutE2eV_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Arrange.Stack(VA.Drawing.Axis.YAxis, Globals.VisioPowerToolsAddIn.addinprefs.SnapUnit);
        }

        private void comboBoxSnapDelta_ValueMemberChanged(object sender, System.EventArgs e)
        {
        }

        private void comboBoxSnapDelta_TextChanged(object sender, System.EventArgs e)
        {
            double d;
            if (double.TryParse(this.comboBoxSnapDelta.Text, out d))
            {
                Globals.VisioPowerToolsAddIn.addinprefs.SnapUnit = d;
            }
        }
    }
}
