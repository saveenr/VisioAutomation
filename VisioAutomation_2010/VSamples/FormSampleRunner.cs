using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using VSamples.Text;

namespace VSamples
{
    public partial class FormSampleRunner : Form
    {
        private readonly List<SampleMethod> _samplemethods = new List<SampleMethod>();
        private readonly Dictionary<string, SampleMethod> _dic = new Dictionary<string, SampleMethod>();

        public FormSampleRunner()
        {
            this.InitializeComponent();

            var methods = new SampleMethods();

            var sm1 = methods.Add(nameof(DeveloperSamples.VisioAutomationNamespacesAndClasses),DeveloperSamples.VisioAutomationNamespacesAndClasses);
            var sm2 = methods.Add(nameof(DeveloperSamples.VisioAutomationNamespaces), DeveloperSamples.VisioAutomationNamespaces);
            var sm3 = methods.Add(nameof(DeveloperSamples.InteropEnumDocumentation), DeveloperSamples.InteropEnumDocumentation);
            var sm4 = methods.Add(nameof(DeveloperSamples.ScriptingDocumentation), DeveloperSamples.ScriptingDocumentation);

            var sm5 = methods.Add(nameof(ConnectorSamples.ConnectorsToBack), ConnectorSamples.ConnectorsToBack);
            var sm6 = methods.Add(nameof(EffectsSamples.GradientTransparencies), EffectsSamples.GradientTransparencies);
            var sm7 = methods.Add(nameof(StencilSamples.DrawGridOfMasters), StencilSamples.DrawGridOfMasters);

            var sm8 = methods.Add(nameof(Layouts.BoxLayout2Samples), Layouts.BoxLayout2Samples.BoxLayout_SimpleCases);
            var sm9 = methods.Add(nameof(Layouts.BoxLayout2Samples), Layouts.BoxLayout2Samples.BoxLayout_TwoLevelGrouping);

            var sm10 = methods.Add(nameof(Layouts.BoxLayoutSamples), Layouts.BoxLayoutSamples.FontCompare);

            var sm11 = methods.Add(nameof(Layouts.ContainerLayoutSamples), Layouts.ContainerLayoutSamples.SimpleContainer);
            var sm12 = methods.Add(nameof(Layouts.DirectedGraphLayoutSamples), Layouts.DirectedGraphLayoutSamples.DirectedGraphViaMsagl);
            var sm13 = methods.Add(nameof(Layouts.DirectedGraphLayoutSamples), Layouts.DirectedGraphLayoutSamples.DirectedGraphViaVisio);
            
            var sm14 = methods.Add(nameof(Layouts.GridLayoutSamples), Layouts.GridLayoutSamples.ColorGrid);
            var sm15 = methods.Add(nameof(Layouts.GridLayoutSamples), Layouts.GridLayoutSamples.ColorGrid);

            var sm16 = methods.Add(nameof(Layouts.TreeLayoutSamples), Layouts.TreeLayoutSamples.TreeWithTwoPassLayoutAndFormatting);

            var sm17 = methods.Add(nameof(CustomPropertySamples), CustomPropertySamples.SetCustomProperties);
            var sm18 = methods.Add(nameof(CustomPropertySamples), CustomPropertySamples.SetCustomProperties);
            var sm19 = methods.Add(nameof(FormsSamples), FormsSamples.MonitorResolutions);
            var sm20 = methods.Add(nameof(PlaygroundSamples), PlaygroundSamples.DrawAllGradients);
            var sm21 = methods.Add(nameof(PlaygroundSamples), PlaygroundSamples.Spirograph);
            var sm22 = methods.Add(nameof(SmartShapeSamples), SmartShapeSamples.ProgressBar);
            var sm23 = methods.Add(nameof(SimpleGeometrySamples), SimpleGeometrySamples.BezierCircle);
            var sm24 = methods.Add(nameof(SimpleGeometrySamples), SimpleGeometrySamples.BezierEllipse);
            var sm25 = methods.Add(nameof(SimpleGeometrySamples), SimpleGeometrySamples.BezierSimple);
            var sm26 = methods.Add(nameof(SimpleGeometrySamples), SimpleGeometrySamples.Nurbs1);
            var sm27 = methods.Add(nameof(SimpleGeometrySamples), SimpleGeometrySamples.Nurbs2);

            var sm29 = methods.Add(nameof(SpecialDocumentSamples), SpecialDocumentSamples.OrgChart);


            var sm30 = methods.Add(nameof(TextMarkpSamples), TextMarkpSamples.TextMarkup11);
            var sm31 = methods.Add(nameof(TextMarkpSamples), TextMarkpSamples.TextMarkup12);
            var sm32 = methods.Add(nameof(TextMarkpSamples), TextMarkpSamples.TextMarkup13);
            var sm33 = methods.Add(nameof(TextMarkpSamples), TextMarkpSamples.TextMarkup14);
            var sm34 = methods.Add(nameof(TextMarkpSamples), TextMarkpSamples.TextMarkup5);
            var sm35 = methods.Add(nameof(TextSamples), TextSamples.NonRotatingText);
            var sm36 = methods.Add(nameof(TextSamples), TextSamples.TextFields);




            var names = new List<string>();


            foreach (var method in methods)
            {
                names.Add(method.Name);
                this._dic[method.Name] = method;
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
                    selectedMethod.Run();
                }
                catch (System.Exception)
                {
                    System.Console.WriteLine("Caught Exception for {0}", selectedMethod.Name);
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