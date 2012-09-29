using System.Linq;
using System.Windows.Forms;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VA = VisioAutomation;

namespace VisioPowerTools
{
    public partial class FormFillDesigner : Form
    {
        public FormFillDesigner()
        {
            InitializeComponent();

            this.update_from_selection();
        }

        private void buttonSetFillGradient_Click(object sender, System.EventArgs e)
        {
            var ss = VisioPowerToolsAddIn.ScriptingSession;
            var selection = ss.Selection.Get();
            if (selection.Count < 1)
            {
                return;
            }

            var format = new VA.Format.ShapeFormatCells();
            format.FillPattern = (int) this.fillGradient1.FillDef.FillPattern;
            format.FillForegnd= VA.Convert.ColorToFormulaRGB(this.fillGradient1.FillDef.ForegroundColor);
            format.FillBkgnd= VA.Convert.ColorToFormulaRGB(this.fillGradient1.FillDef.BackgroundColor);
            format.FillForegndTrans= this.fillGradient1.FillDef.ForegroundTransparency/100.0;
            format.FillBkgndTrans= this.fillGradient1.FillDef.BackgroundTransparency / 100.0;
            format.ShdwPattern= (int) this.fillGradient1.ShadowDef.FillPattern;
            format.ShdwForegnd= VA.Convert.ColorToFormulaRGB(this.fillGradient1.ShadowDef.ForegroundColor);
            format.ShdwBkgnd = VA.Convert.ColorToFormulaRGB(this.fillGradient1.ShadowDef.BackgroundColor);
            format.ShdwForegndTrans = this.fillGradient1.ShadowDef.ForegroundTransparency/100.0;
            format.ShdwBkgndTrans= this.fillGradient1.ShadowDef.BackgroundTransparency/100.0;
            format.ShapeShdwOffsetX = 0.0;
            format.ShapeShdwOffsetY = 0.0;
            format.ShapeShdwScaleFactor = 1.0;
            format.ShapeShdwType= 5;


            var update = new VA.ShapeSheet.Update();
            var shapes = ss.Selection.EnumShapes().ToList();
            var shapeids = shapes.Select(s => s.ID).ToList();

            foreach (int shapeid in shapeids)
            {
                format.Apply(update, (short)shapeid);
            }

            update.Execute(ss.VisioApplication.ActivePage);    
        }

        private void buttonUpdateFill_Click(object sender, System.EventArgs e)
        {
            this.update_from_selection();
        }

        private void update_from_selection()
        {
            var app = VisioPowerToolsAddIn.ScriptingSession;

            if (!app.Selection.HasShapes())
            {
                return;
            }

            var application = app.VisioApplication;
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            var s1 = selection[1];
            var doc = application.ActiveDocument;

            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_fg = query.AddColumn(VA.ShapeSheet.SRCConstants.FillForegnd);
            var col_bg = query.AddColumn(VA.ShapeSheet.SRCConstants.FillBkgnd);
            var col_fgtrans = query.AddColumn(VA.ShapeSheet.SRCConstants.FillForegndTrans);
            var col_bgtrans = query.AddColumn(VA.ShapeSheet.SRCConstants.FillBkgndTrans);
            var col_fillpat = query.AddColumn(VA.ShapeSheet.SRCConstants.FillPattern);
            var col_sfg = query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwForegnd);
            var col_sbg = query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwBkgnd);
            var col_sfgtrans = query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwForegndTrans);
            var col_bfgtrans = query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwBkgndTrans);
            var col_spat = query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwPattern);

            var table = query.GetResults<double>(s1);
            var colors = doc.Colors;

            var row = table[0];
            this.fillGradient1.FillDef.ForegroundColor = (System.Drawing.Color) colors[(int)row[col_fg]].ToColorRGB();
            this.fillGradient1.FillDef.BackgroundColor = (System.Drawing.Color)colors[(int)row[col_bg]].ToColorRGB();
            this.fillGradient1.FillDef.ForegroundTransparency = (int)(100.0 * row[col_fgtrans]);
            this.fillGradient1.FillDef.BackgroundTransparency = (int)(100.0 * row[col_bgtrans]);
            this.fillGradient1.FillDef.FillPattern = (VA.UI.FillPattern)(int)row[col_fillpat];

            this.fillGradient1.ShadowDef.ForegroundColor = (System.Drawing.Color)colors[(int)row[col_sfg]].ToColorRGB();
            this.fillGradient1.ShadowDef.BackgroundColor = (System.Drawing.Color)colors[(int)row[col_sbg]].ToColorRGB();
            this.fillGradient1.ShadowDef.ForegroundTransparency = (int)(100.0 * row[col_sfgtrans]);
            this.fillGradient1.ShadowDef.BackgroundTransparency = (int)(100.0 * row[col_bfgtrans]);
            this.fillGradient1.ShadowDef.FillPattern = (VA.UI.FillPattern)((int)row[col_spat]);
        }
    }
}