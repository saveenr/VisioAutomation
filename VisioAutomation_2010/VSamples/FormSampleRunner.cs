using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace VSamples
{
    public partial class FormSampleRunner : Form
    {
        private readonly List<SampleMethod> _samplemethods = new List<SampleMethod>();
        private readonly Dictionary<string, SampleMethod> _dic = new Dictionary<string, SampleMethod>();

        public FormSampleRunner()
        {
            this.InitializeComponent();

            var sm1 = new SampleMethod(nameof(DeveloperSamples.VisioAutomationNamespacesAndClasses),DeveloperSamples.VisioAutomationNamespacesAndClasses);
            var sm2 = new SampleMethod(nameof(DeveloperSamples.VisioAutomationNamespaces), DeveloperSamples.VisioAutomationNamespaces);
            var sm3 = new SampleMethod(nameof(DeveloperSamples.InteropEnumDocumentation), DeveloperSamples.InteropEnumDocumentation);
            var sm4 = new SampleMethod(nameof(DeveloperSamples.ScriptingDocumentation), DeveloperSamples.ScriptingDocumentation);

            var methods = new List<SampleMethod>();
            methods.Add(sm1);
            methods.Add(sm2);
            methods.Add(sm3);
            methods.Add(sm4);

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