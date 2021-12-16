using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace VSamples
{
    public partial class FormSampleRunner : Form
    {
        private readonly Dictionary<string, SampleMethod> dic_name_to_sample = new Dictionary<string, SampleMethod>();

        public FormSampleRunner()
        {
            this.InitializeComponent();

            var samples = AddSampleMethods();
            this.dic_name_to_sample = samples.ToDictionary(x => x.Name, x => x);
            var samplenames = samples.Select(x => x.Name).OrderBy(i => i).ToList();

            var prev_names = this.GetPreviouslySelectedSamples();

            foreach (var name in samplenames)
            {
                bool ischecked = prev_names.Contains(name);
                this.checkedListBox1.Items.Add(name, ischecked);
            }
        }

        private static List<SampleMethod> AddSampleMethods()
        {
            var methods = new List<SampleMethod>();


            methods.Add(new Samples.Developer.DiagramVAClasses());
            methods.Add(new Samples.Developer.DiagramVANamespaces());
            methods.Add(new Samples.Developer.DocumentScriptingAPI());
            methods.Add(new Samples.Developer.DocumentVisioInterop());
            methods.Add(new Samples.Layouts.BoxLayoutSimple());
            methods.Add(new Samples.Layouts.BoxLayoutTwoLevelGrouping());
            methods.Add(new Samples.Layouts.ColorGrid());
            methods.Add(new Samples.Layouts.CompareFonts());
            methods.Add(new Samples.Layouts.Container1());
            methods.Add(new Samples.Layouts.DirectedGraphViaMsagl());
            methods.Add(new Samples.Layouts.DirectedGraphViaVisio());
            methods.Add(new Samples.Layouts.TreeWithTwoPassLayoutAndFormatting());
            methods.Add(new Samples.Misc.AllGradients());
            methods.Add(new Samples.Misc.BezierCircle());
            methods.Add(new Samples.Misc.BezierEllipse());
            methods.Add(new Samples.Misc.BezierSimple());
            methods.Add(new Samples.Misc.GradientTransparencies());
            methods.Add(new Samples.Misc.GridOfMasters());
            methods.Add(new Samples.Misc.MonitorResolutions());
            methods.Add(new Samples.Misc.Nurbs1());
            methods.Add(new Samples.Misc.Nurbs2());
            methods.Add(new Samples.Misc.ProgressBar());
            methods.Add(new Samples.Misc.OrgChart1());
            methods.Add(new Samples.Misc.PathAnalysisSamples());
            methods.Add(new Samples.Misc.SendConnectorsToBack());
            methods.Add(new Samples.Misc.SetCustomProperties());
            methods.Add(new Samples.Misc.Spirograph());
            methods.Add(new Samples.Text.NonRotatingText());
            methods.Add(new Samples.Text.TextFields());
            methods.Add(new Samples.Text.TextMarkup1());
            methods.Add(new Samples.Text.TextMarkup2());
            methods.Add(new Samples.Text.TextMarkup3());
            methods.Add(new Samples.Text.TextMarkup4());
            methods.Add(new Samples.Text.TextMarkup5());


            return methods;
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

            var selected_methods = selected_names.Select(n => this.dic_name_to_sample[n]).ToList();

            foreach (var selectedMethod in selected_methods)
            {
                try
                {
                    selectedMethod.Run();
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