using VisioAutomation.Extensions;
using VisioAutomation.Models.Layouts.Grid;
using VisioAutomation.ShapeSheet.Writers;

namespace VSamples.Samples.Misc
{
    public class AllGradients : SampleMethodBase
    {
        public override void RunSample()
        {
            var app = SampleEnvironment.Application;
            var docs = app.Documents;
            var stencil = docs.OpenStencil("basic_u.vss");
            var master = stencil.Masters["Rectangle"];
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            int num_cols = 7;
            int num_rows = 7;

            var page_size = new VisioAutomation.Core.Size(5, 5);
            SampleEnvironment.SetPageSize(page, page_size);

            var lowerleft = new VisioAutomation.Core.Point(0, 0);
            var actual_page_size = SampleEnvironment.GetPageSize(page);
            var page_rect = new VisioAutomation.Core.Rectangle(lowerleft, actual_page_size);

            var layout = new GridLayout(num_cols, num_rows, new VisioAutomation.Core.Size(1, 1), master);
            layout.RowDirection = RowDirection.TopToBottom;
            layout.Origin = page_rect.UpperLeft;
            layout.CellSpacing = new VisioAutomation.Core.Size(0, 0);
            layout.PerformLayout();

            int max_grad_id = 40;
            int n = 0;

            foreach (var node in layout.Nodes)
            {
                int grad_id = n % max_grad_id;
                node.Text = grad_id.ToString();
                n++;
            }

            layout.Render(page);

            var color1 = new VisioAutomation.Models.Color.ColorRgb(0xffdddd);
            var color2 = new VisioAutomation.Models.Color.ColorRgb(0x00ffff);

            var format = new VisioAutomation.Shapes.ShapeFormatCells();

            var writer = new SidSrcWriter();

            string color1_formula = color1.ToFormula();
            string color2_formula = color2.ToFormula();

            n = 0;

            foreach (var node in layout.Nodes)
            {
                short shapeid = node.ShapeID;
                int grad_id = n % max_grad_id;

                format.FillPattern = grad_id;
                format.FillForeground = color1_formula;
                format.FillBackground = color2_formula;
                format.LinePattern = 0;
                format.LineWeight = 0;

                writer.SetValues(shapeid, format);

                n++;
            }

            writer.Commit(page, VisioAutomation.Core.CellValueType.Formula);

            var bordersize = new VisioAutomation.Core.Size(1, 1);
            page.ResizeToFitContents(bordersize);
        }
    }
}