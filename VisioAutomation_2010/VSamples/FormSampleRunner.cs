using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using VSamples.Samples.Developer;
using VSamples.Samples.Layouts;
using VSamples.Samples.Misc;
using VSamples.Samples.Text;
using ProgressBar = VSamples.Samples.Misc.ProgressBar;

namespace VSamples
{
    public partial class FormSampleRunner : Form
    {
        private readonly Dictionary<string, SampleMethodBase> _dic = new Dictionary<string, SampleMethodBase>();

        public FormSampleRunner()
        {
            this.InitializeComponent();

            var methods = new SampleMethods();

            var sm1 = methods.AddEx(new DocumentScriptingAPI());
            var sm2 = methods.AddEx(new DocumentVisioInterop());
            var sm3 = methods.AddEx(new DiagramVANamespaces());
            var sm4 = methods.AddEx(new DiagramVAClasses());

            var sm5 = methods.AddEx(new SendConnectorsToBack());
            var sm6 = methods.AddEx(new GradientTransparencies());
            var sm7 = methods.AddEx(new GridOfMasters());

            var sm8 = methods.AddEx(new BoxLayoutSimple());
            var sm9 = methods.AddEx(new BoxLayoutTwoLevelGrouping());

            var sm10 = methods.AddEx(new CompareFonts());

            var sm11 = methods.AddEx(new Container1());
            var sm12 = methods.AddEx(new DirectedGraphViaMsagl());
            var sm13 = methods.AddEx(new DirectedGraphViaVisio());
            
            var sm14 = methods.AddEx(new ColorGrid());

            var sm16 = methods.AddEx(new TreeWithTwoPassLayoutAndFormatting());

            var sm17 = methods.AddEx(new SetCustomProperties());

            var sm19 = methods.AddEx(new MonitorResolutions());
            var sm20 = methods.AddEx(new AllGradients());
            var sm21 = methods.AddEx(new Spirograph());
            var sm22 = methods.AddEx(new ProgressBar());
            var sm23 = methods.AddEx(new BezierCircle());
            var sm24 = methods.AddEx(new BezierEllipse());
            var sm25 = methods.AddEx(new BezierSimple());
            var sm26 = methods.AddEx(new Nurbs1());
            var sm27 = methods.AddEx(new Nurbs2());

            var sm29 = methods.AddEx(new OrgChart1());


            var sm30 = methods.AddEx(new TextMarkup1());
            var sm31 = methods.AddEx(new TextMarkup2());
            var sm32 = methods.AddEx(new TextMarkup3());
            var sm33 = methods.AddEx(new TextMarkup4());
            var sm34 = methods.AddEx(new TextMarkup5());

            var sm35 = methods.AddEx(new NonRotatingText());
            var sm36 = methods.AddEx(new TextFields());




            var names = new List<string>();


            foreach (var method in methods)
            {
                names.Add(method.GetName());
                this._dic[method.GetName()] = method;
            }

            var prev_names = this.GetPreviouslySelectedSamples();

            foreach (var name in names)
            {
                bool ischecked = prev_names.Contains(name);
                this.checkedListBox1.Items.Add(name, ischecked);

            }

            const bool autorun = false;
            if (autorun)
            {
            // this.RunSelectedSamples();
            }
        }

        private HashSet<string> GetPreviouslySelectedSamples()
        {
            var prev_names_str = Properties.Settings.Default.SelectedSamples ?? "";
            return new HashSet<string>(prev_names_str.Split('|'));
        }

        private void SaveSelectedNames()
        {
            var selected_names = this.GetSelectedNames();
            Properties.Settings.Default.SelectedSamples = string.Join("|", selected_names);
            Properties.Settings.Default.Save();
        }

        private void buttonRun_Click(object sender, System.EventArgs e)
        {
            this.RunSelectedSamples();
        }

        private void RunSelectedSamples()
        {
            var selected_names = this.GetSelectedNames();

            this.SaveSelectedNames();

            var selected_methods = selected_names.Select(n => this._dic[n]).ToList();

            foreach (var selectedMethod in selected_methods)
            {
                try
                {
                    selectedMethod.RunSample();
                }
                catch (System.Exception)
                {
                    System.Console.WriteLine("Caught Exception for {0}", selectedMethod.GetType());
                    break;
                }
            }
        }

        private List<string> GetSelectedNames()
        {
            var selected_names = new List<string>();
            foreach (var item in this.checkedListBox1.CheckedItems)
            {
                selected_names.Add((string) item);
            }
            return selected_names;
        }

        private void buttonSelectAll_Click(object sender, System.EventArgs e)
        {
            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {
                this.checkedListBox1.SetItemCheckState(i, CheckState.Checked);
            }
        }

        private void buttonSelectNone_Click(object sender, System.EventArgs e)
        {
            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {
                this.checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
            }
        }

        private void FormSampleRunner_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        private void FormSampleRunner_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.SaveSelectedNames();
        }
    }
}