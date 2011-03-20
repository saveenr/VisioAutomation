using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomationDevTool
{
    public partial class FormVADevTool : Form
    {
        private IVisio.Application app;

        public FormVADevTool()
        {
            InitializeComponent();
        }

        private void buttonLaunchVisio_Click(object sender, EventArgs e)
        {
            this.app = new IVisio.Application();
        }

        private void buttonGetUnitTestCode_Click(object sender, EventArgs e)
        {
            if (this.app==null)
            {
                return;
            }

            var docs = app.Documents;

            if (docs.Count<1)
            {
                return; ;
            }

            var page = app.ActivePage;
            if (page==null)
            {
                return;
            }

            var pageshapes = page.Shapes.AsEnumerable().ToList();

            var shapeids = pageshapes.Select( s=>s.ID).ToList();
            var xforms = VA.Layout.LayoutHelper.GetXForm(page, shapeids);


            var lines = new List<string>();
            foreach (int i in Enumerable.Range(0,shapeids.Count))
            {
                lines.Add("");
                var xform = xforms[i];
                var shape = pageshapes[i];
                string line = string.Format(@"// {0} ", shape.NameU);
                lines.Add(line);
                string shape_var = string.Format("test_shape_{0}",i);
                line = string.Format(@"var {0} = page.Shapes[""{1}""];",shape_var, shape.NameU);
                lines.Add(line);

                var cellnames = new[] { "PinX", "PinY", "LocPinX", "LocPinY", "Width", "Height" , "Angle"};
                foreach (var cellname in cellnames)
                {
                    line = string.Format(@"Assert.AreEqual(  ""{0}"" , {1}.CellsU[""{2}""].Formula  ); ", shape.CellsU[cellname].Formula, shape_var, cellname);
                    lines.Add(line);
                   
                }

                line = string.Format(@"Assert.AreEqual(  ""{0}"" , {1}.Text ); ", shape.Text, shape_var);
                lines.Add(line);
            }


            var form = new FormTextWindow();
            form.SetText(lines);

            form.ShowDialog();

        }
    }
}
