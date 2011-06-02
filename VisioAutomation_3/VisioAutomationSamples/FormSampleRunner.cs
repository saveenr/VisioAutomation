using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace VisioAutomationSamples
{
    public partial class FormSampleRunner : Form
    {
        private List<SampleMethod> samplemethods = new List<SampleMethod>();
        private Dictionary<string, SampleMethod> dic = new Dictionary<string, SampleMethod>();

        public FormSampleRunner()
        {
            InitializeComponent();

            var all_types = typeof (Program).Assembly.GetExportedTypes();
            var public_sample_classes = all_types
                .Where(t => t.IsPublic)
                .Where(t => t.IsClass)
                .Where(t => t.Name.Contains("Sample"))
                .OrderBy(t => t.Name);

            var names = new List<string>();
            foreach (var t in public_sample_classes)
            {
                var methods = t.GetMethods()
                    .Where(m => m.IsPublic)
                    .Where(m => m.IsStatic)
                    .Where(m => m.GetParameters().Count() == 0)
                    .OrderBy(m => m.Name);

                foreach (var m in methods)
                {
                    string name = string.Format("{0} / {1}", t.Name, m.Name);
                    names.Add(name);

                    var item = new SampleMethod();
                    item.Name = name;
                    item.Method = m;

                    samplemethods.Add(item);

                    dic[name] = item;
                }
            }

            var prev_names = GetPreviouslySelectedSamples();

            foreach (var name in names)
            {
                bool ischecked = prev_names.Contains(name);
                this.checkedListBox1.Items.Add(name, ischecked);
            }
        }

        private HashSet<string> GetPreviouslySelectedSamples()
        {
            var prev_names_str = Properties.Settings.Default.SelectedSamples;
            if (prev_names_str == null)
            {
                prev_names_str = "";
            }
            return new HashSet<string>(prev_names_str.Split(new char[] {'|'}));
        }

        private void SaveSelectedNames()
        {
            var selected_names = GetSelectedNames();
            Properties.Settings.Default.SelectedSamples = string.Join("|", selected_names);
            Properties.Settings.Default.Save();
        }

        private void buttonRun_Click(object sender, EventArgs e)
        {
            var selected_names = GetSelectedNames();

            this.SaveSelectedNames();

            var selected_methods = selected_names.Select(n => dic[n]).ToList();

            foreach (var selectedMethod in selected_methods)
            {
                try
                {
                    selectedMethod.Run();
                }
                catch (Exception)
                {
                    Console.WriteLine("Caught Exception for {0}", selectedMethod.Name);
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

        private void buttonSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {
                this.checkedListBox1.SetItemCheckState(i, CheckState.Checked);
            }
        }

        private void buttonSelectNone_Click(object sender, EventArgs e)
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