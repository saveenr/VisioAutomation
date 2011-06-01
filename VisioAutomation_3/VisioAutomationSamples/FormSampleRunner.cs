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
        List<SampleMethod> samplemethods = new List<SampleMethod>();
        Dictionary<string,SampleMethod> dic = new Dictionary<string, SampleMethod>();

        public FormSampleRunner()
        {
            InitializeComponent();

            var all_types = typeof (Program).Assembly.GetExportedTypes();
            var public_sample_classes = all_types
                .Where(t => t.IsPublic)
                .Where(t => t.IsClass)
                .Where(t=>t.Name.Contains("Sample"))
                .OrderBy(t=>t.Name);

            var names = new List<string>();
            foreach (var t in public_sample_classes)
            {
                var methods = t.GetMethods()
                    .Where(m => m.IsPublic)
                    .Where(m => m.IsStatic)
                    .Where(m => m.GetParameters().Count() == 0)
                    .OrderBy(m=>m.Name);

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

            foreach (var name in names)
            {
                this.checkedListBox1.Items.Add(name);
            }


        }

        private void buttonRun_Click(object sender, EventArgs e)
        {
            var selected_names = new List<string>();
            foreach (var item in this.checkedListBox1.CheckedItems)
            {
                selected_names.Add( (string) item);
            }

            var selected_methods = selected_names.Select(n => dic[n]).ToList();

            foreach (var selectedMethod in selected_methods)
            {
                try
                {
                    selectedMethod.Run();
                }
                catch (Exception)
                {
                    System.Windows.Forms.MessageBox.Show("Caught Exception");
                    break;
                }
            }
        }

        private void buttonSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {
                this.checkedListBox1.SetItemCheckState(i,CheckState.Checked);
            }
        }

        private void buttonSelectNone_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {
                this.checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
            }
        }
    }

    public class SampleMethod
    {
        public string Name;
        public System.Reflection.MethodInfo Method;

        public void Run()
        {
            this.Method.Invoke(null, null);
        }

    }
}
