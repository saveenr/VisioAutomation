using System.Windows.Forms;

namespace VisioPowerTools2010
{
    public partial class FormDirectedGraph : Form
    {
        public FormDirectedGraph()
        {
            this.InitializeComponent();
        }

        public string GraphText
        {
            get { return this.textBox1.Text; }
        }
    }
}
