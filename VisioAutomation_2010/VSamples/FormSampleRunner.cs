using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using VSamples.Text;

namespace VSamples
{
    public partial class FormSampleRunner : Form
    {
        private readonly Dictionary<string, SampleMethodBase> _dic = new Dictionary<string, SampleMethodBase>();

        public FormSampleRunner()
        {
            this.InitializeComponent();

            var methods = new SampleMethods();

            var sm1 = methods.AddEx(new ScriptingDocumentationXX());
            var sm2 = methods.AddEx(new InteropEnumDocumentationX());
            var sm3 = methods.AddEx(new VisioAutomationNamespacesX());
            var sm4 = methods.AddEx(new VisioAutomationNamespacesAndClassesX());

            var sm5 = methods.AddEx(new ConnectorsToBackX());
            var sm6 = methods.AddEx(new GradientTransparenciesX());
            var sm7 = methods.AddEx(new DrawGridOfMastersX());

            var sm8 = methods.AddEx(new Layouts.BoxLayout_SimpleCasesx());
            var sm9 = methods.AddEx(new Layouts.BoxLayout_TwoLevelGroupingx());

            var sm10 = methods.AddEx(new Layouts.FontCompareX());

            var sm11 = methods.AddEx(new Layouts.SimpleContainerX());
            var sm12 = methods.AddEx(new Layouts.DirectedGraphViaMsaglX());
            var sm13 = methods.AddEx(new Layouts.DirectedGraphViaVisioX());
            
            var sm14 = methods.AddEx(new Layouts.ColorGridX());

            var sm16 = methods.AddEx(new Layouts.TreeWithTwoPassLayoutAndFormattingX());

            var sm17 = methods.AddEx(new SetCustomPropertiesX());

            var sm19 = methods.AddEx(new MonitorResolutionsX());
            var sm20 = methods.AddEx(new DrawAllGradientsX());
            var sm21 = methods.AddEx(new SpirographX());
            var sm22 = methods.AddEx(new ProgressBarX());
            var sm23 = methods.AddEx(new BezierCircleX());
            var sm24 = methods.AddEx(new BezierEllipseX());
            var sm25 = methods.AddEx(new BezierSimpleX());
            var sm26 = methods.AddEx(new Nurbs1X());
            var sm27 = methods.AddEx(new Nurbs2X());

            var sm29 = methods.AddEx(new OrgChartX());


            var sm30 = methods.AddEx(new TextMarkpSamples1());
            var sm31 = methods.AddEx(new TextMarkpSamples2());
            var sm32 = methods.AddEx(new TextMarkpSamples3());
            var sm33 = methods.AddEx(new TextMarkpSamples4());
            var sm34 = methods.AddEx(new TextMarkpSamples5());

            var sm35 = methods.AddEx(new NonRotatingTextX());
            var sm36 = methods.AddEx(new TextFieldsX());




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