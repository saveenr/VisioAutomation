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

        private void buttonSet2ColorGlow_Click(object sender, System.EventArgs e)
        {
            var vis = Globals.VisioPowerToolsAddIn.Application;
            var active_window = vis.ActiveWindow;
            var selection = active_window.Selection;
            if (selection.Count < 1)
            {
                return;
            }

            var upper_glow_color = new VA.Drawing.ColorRGB(this.uC2ColorGlow1.UpperColor);
            var lower_glow_color = new VA.Drawing.ColorRGB(this.uC2ColorGlow1.LowerColor);
            double uppertrans = this.uC2ColorGlow1.UpperTransparency/100.0;
            double lowertrans = this.uC2ColorGlow1.LowerTransparency/100.0;
            double scale = this.uC2ColorGlow1.GlowSize/100.0;

            var fildef = new VA.Effects.TwoColorGlow();
            fildef.TopColor = upper_glow_color;
            fildef.BottomColor = lower_glow_color;
            fildef.TopTransparency = uppertrans;
            fildef.BottomTransparency = lowertrans;
            fildef.Scale = scale;

            var fmt = fildef.GetFormat();

            VisioPowerToolsAddIn.ScriptingSession.Format.SetFormat(fmt);
        }

        private void buttonSet3PointFill_Click(object sender, System.EventArgs e)
        {
            var vis = Globals.VisioPowerToolsAddIn.Application;
            var active_window = vis.ActiveWindow;
            var selection = active_window.Selection;
            if (selection.Count < 1)
            {
                return;
            }

            var filldef = new VA.Effects.ThreePointGradientFill();
            filldef.Direction = this.uC3PointFill1.Direction;
            filldef.SideColor = new VA.Drawing.ColorRGB(this.uC3PointFill1.EdgeColor);
            filldef.SideTransparency = 0.0;
            filldef.Corner1Color = new VA.Drawing.ColorRGB(this.uC3PointFill1.Corner1Color);
            filldef.Corner1Transparency = 0.0;
            filldef.Corner2Color = new VA.Drawing.ColorRGB(this.uC3PointFill1.Corner2Color);
            filldef.Corner2Transparency = 0.0;
            var fmt = filldef.GetFormat();

            VisioPowerToolsAddIn.ScriptingSession.Format.SetFormat(fmt);
        }

        private void buttonSetFillGradient_Click(object sender, System.EventArgs e)
        {
            var ss = VisioPowerToolsAddIn.ScriptingSession;
            var selection = ss.Selection.GetSelection();
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


            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            var shapes = ss.Selection.EnumSelectedShapes().ToList();
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

            if (!app.Selection.HasSelectedShapes())
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

            var row = table.Rows[0];
            this.fillGradient1.FillDef.ForegroundColor = (System.Drawing.Color) colors[(int)row[col_fg]].ToColorRGB();
            this.fillGradient1.FillDef.BackgroundColor = (System.Drawing.Color)colors[(int)row[col_bg]].ToColorRGB();
            this.fillGradient1.FillDef.ForegroundTransparency = (int)(100.0 * row[col_fgtrans]);
            this.fillGradient1.FillDef.BackgroundTransparency = (int)(100.0 * row[col_bgtrans]);
            this.fillGradient1.FillDef.FillPattern = (VA.Format.FillPattern)(int)row[col_fillpat];

            this.fillGradient1.ShadowDef.ForegroundColor = (System.Drawing.Color)colors[(int)row[col_sfg]].ToColorRGB();
            this.fillGradient1.ShadowDef.BackgroundColor = (System.Drawing.Color)colors[(int)row[col_sbg]].ToColorRGB();
            this.fillGradient1.ShadowDef.ForegroundTransparency = (int)(100.0 * row[col_sfgtrans]);
            this.fillGradient1.ShadowDef.BackgroundTransparency = (int)(100.0 * row[col_bfgtrans]);
            this.fillGradient1.ShadowDef.FillPattern = (VA.Format.FillPattern)((int)row[col_spat]);
        }
    }
}