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

            var sm1 = methods.Add(nameof(ScriptingDocumentationXX.ScriptingDocumentation), ScriptingDocumentationXX.ScriptingDocumentation);
            var sm2 = methods.Add(nameof(InteropEnumDocumentationX.InteropEnumDocumentation), InteropEnumDocumentationX.InteropEnumDocumentation);
            var sm3 = methods.Add(nameof(VisioAutomationNamespacesX.VisioAutomationNamespaces), VisioAutomationNamespacesX.VisioAutomationNamespaces);
            var sm4 = methods.Add(nameof(VisioAutomationNamespacesAndClassesX.VisioAutomationNamespacesAndClasses), VisioAutomationNamespacesAndClassesX.VisioAutomationNamespacesAndClasses);

            var sm5 = methods.Add(nameof(ConnectorsToBackX.ConnectorsToBack), ConnectorsToBackX.ConnectorsToBack);
            var sm6 = methods.Add(nameof(GradientTransparenciesX.GradientTransparencies), GradientTransparenciesX.GradientTransparencies);
            var sm7 = methods.Add(nameof(DrawGridOfMastersX.DrawGridOfMasters), DrawGridOfMastersX.DrawGridOfMasters);

            var sm8 = methods.Add(nameof(Layouts.BoxLayout_SimpleCasesx), Layouts.BoxLayout_SimpleCasesx.BoxLayout_SimpleCases);
            var sm9 = methods.Add(nameof(Layouts.BoxLayout_TwoLevelGroupingx), Layouts.BoxLayout_TwoLevelGroupingx.BoxLayout_TwoLevelGrouping);

            var sm10 = methods.Add(nameof(Layouts.FontCompareX), Layouts.FontCompareX.FontCompare);

            var sm11 = methods.Add(nameof(Layouts.SimpleContainerX), Layouts.SimpleContainerX.SimpleContainer);
            var sm12 = methods.Add(nameof(Layouts.DirectedGraphViaMsaglX), Layouts.DirectedGraphViaMsaglX.DirectedGraphViaMsagl);
            var sm13 = methods.Add(nameof(Layouts.DirectedGraphViaVisioX), Layouts.DirectedGraphViaVisioX.DirectedGraphViaVisio);
            
            var sm14 = methods.Add(nameof(Layouts.ColorGridX), Layouts.ColorGridX.ColorGrid);

            var sm16 = methods.Add(nameof(Layouts.TreeWithTwoPassLayoutAndFormattingX), Layouts.TreeWithTwoPassLayoutAndFormattingX.TreeWithTwoPassLayoutAndFormatting);

            var sm17 = methods.Add(nameof(SetCustomPropertiesX), SetCustomPropertiesX.SetCustomProperties);

            var sm19 = methods.Add(nameof(MonitorResolutionsX), MonitorResolutionsX.MonitorResolutions);
            var sm20 = methods.Add(nameof(DrawAllGradientsX), DrawAllGradientsX.DrawAllGradients);
            var sm21 = methods.Add(nameof(SpirographX), SpirographX.Spirograph);
            var sm22 = methods.Add(nameof(ProgressBarX), ProgressBarX.ProgressBar);
            var sm23 = methods.Add(nameof(BezierCircleX), BezierCircleX.BezierCircle);
            var sm24 = methods.Add(nameof(BezierEllipseX), BezierEllipseX.BezierEllipse);
            var sm25 = methods.Add(nameof(BezierSimpleX), BezierSimpleX.BezierSimple);
            var sm26 = methods.Add(nameof(Nurbs1X), Nurbs1X.Nurbs1);
            var sm27 = methods.Add(nameof(Nurbs2X), Nurbs2X.Nurbs2);

            var sm29 = methods.Add(nameof(OrgChartX), OrgChartX.OrgChart);


            var sm30 = methods.Add(nameof(TextMarkpSamples1), TextMarkpSamples1.TextMarkup11);
            var sm31 = methods.Add(nameof(TextMarkpSamples2), TextMarkpSamples2.TextMarkup12);
            var sm32 = methods.Add(nameof(TextMarkpSamples3), TextMarkpSamples3.TextMarkup13);
            var sm33 = methods.Add(nameof(TextMarkpSamples4), TextMarkpSamples4.TextMarkup14);
            var sm34 = methods.Add(nameof(TextMarkpSamples5), TextMarkpSamples5.TextMarkup5);

            var sm35 = methods.Add(nameof(NonRotatingTextX), NonRotatingTextX.NonRotatingText);
            var sm36 = methods.Add(nameof(TextFieldsX), TextFieldsX.TextFields);




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