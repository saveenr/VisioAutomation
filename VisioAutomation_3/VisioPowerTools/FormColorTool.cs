using System.Windows.Forms;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet;
using VAS = VisioAutomation.Scripting;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerTools
{
    public partial class FormColorTool : Form
    {
        public class ColorItem
        {
            public string Name;
            public VA.ShapeSheet.SRC SRC;
            public VA.ShapeSheet.FormulaLiteral Formula;

            public ColorItem(string name, VA.ShapeSheet.SRC src)
            {
                this.Name = name;
                this.SRC = src;
                this.Formula = null;
            }
        }

        public class ColorCells
        {
            public ColorItem FillBackgroundColor = new ColorItem("Fill Background", VA.ShapeSheet.SRCConstants.FillBkgnd);
            public ColorItem FillForegroundColor = new ColorItem("Fill Foreground", VA.ShapeSheet.SRCConstants.FillForegnd);
            public ColorItem LineColor = new ColorItem("Line Color", VA.ShapeSheet.SRCConstants.LineColor);
            public ColorItem ShadowBackgroundColor = new ColorItem("Shadow Background", VA.ShapeSheet.SRCConstants.ShdwBkgnd);

            public ColorItem ShadowForegroundColor = new ColorItem("Shadow Foreground",
                                                                   VA.ShapeSheet.SRCConstants.ShdwForegnd);

            public ColorItem Character = new ColorItem("Character", VA.ShapeSheet.SRCConstants.Char_Color);
        }

        public ColorCells Colors = new ColorCells();

        public FormColorTool()
        {
            InitializeComponent();
        }

        private void buttonRead_Click(object sender, System.EventArgs e)
        {
            var scriptingsession = VisioPowerToolsAddIn.ScriptingSession;
            if (!scriptingsession.HasSelectedShapes())
            {
                return;
            }

            var application = scriptingsession.Application;
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            var s1 = selection[1];
            var doc = application.ActiveDocument;
            var doc_colors = doc.Colors;

            var rgbcolors = ShapeFormatHelper.GetColorsFromShape(s1);

            this.colorSelectorSmallFillFg.Color =
                (System.Drawing.Color) doc_colors[rgbcolors.FillForegroundColor.Result].ToColorRGB();
            this.colorSelectorSmallFillBg.Color =
                (System.Drawing.Color) doc_colors[rgbcolors.FillBackgroundColor.Result].ToColorRGB();
            this.colorSelectorSmallShdFg.Color =
                (System.Drawing.Color) doc_colors[rgbcolors.ShadowForegroundColor.Result].ToColorRGB();
            this.colorSelectorSmallShdBg.Color =
                (System.Drawing.Color) doc_colors[rgbcolors.ShadowBackgroundColor.Result].ToColorRGB();
            this.colorSelectorSmallLine.Color =
                (System.Drawing.Color) doc_colors[rgbcolors.LineColor.Result].ToColorRGB();
        }

        private void buttonApply_Click(object sender, System.EventArgs e)
        {
            var scriptingsession = VisioPowerToolsAddIn.ScriptingSession;

            this.Colors.FillForegroundColor.Formula =
                VA.Convert.ColorToFormulaRGB(this.colorSelectorSmallFillFg.Color);
            this.Colors.FillBackgroundColor.Formula =
                VA.Convert.ColorToFormulaRGB(this.colorSelectorSmallFillBg.Color);
            this.Colors.ShadowForegroundColor.Formula =
                VA.Convert.ColorToFormulaRGB(this.colorSelectorSmallShdFg.Color);
            this.Colors.ShadowBackgroundColor.Formula =
                VA.Convert.ColorToFormulaRGB(this.colorSelectorSmallShdBg.Color);
            this.Colors.LineColor.Formula =
                VA.Convert.ColorToFormulaRGB(this.colorSelectorSmallText.Color);

            var cells = new[]
                            {
                                this.Colors.FillBackgroundColor,
                                this.Colors.FillBackgroundColor,
                                this.Colors.ShadowForegroundColor,
                                this.Colors.ShadowBackgroundColor,
                                this.Colors.LineColor
                            };
            scriptingsession.ShapeSheet.SetFormulas(cells,
                                         i => i.Formula.HasValue,
                                         i => i.SRC,
                                         i => i.Formula.ToString());
        }
    }

    internal static class ShapeFormatHelper
    {
        private static VA.Format.ShapeFormatQuery _format_query = new VA.Format.ShapeFormatQuery();

        public static ShapeColors GetColorsFromShape(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var r = _format_query.GetFormulasAndResults<double>(shape);

            var shapecolors = new ShapeFormatHelper.ShapeColors();
            var fgord = _format_query.FillForegnd;
            var bgord = _format_query.FillBkgnd;
            var sfgord = _format_query.ShdwForegnd;
            var bggord = _format_query.ShdwBkgnd;
            var lcord = _format_query.LineColor;
            var ccord = _format_query.CharColor;

            var tablef = r.Formulas;
            var tabler = r.Results;

            var fr = tablef.Rows[0];
            var rr = tabler.Rows[0];
            shapecolors.FillForegroundColor = new CellData<int>(fr[fgord], (int) rr[fgord]);
            shapecolors.FillBackgroundColor = new CellData<int>(fr[bgord], (int) rr[bgord]);
            shapecolors.ShadowForegroundColor = new CellData<int>(fr[sfgord], (int) rr[sfgord]);
            shapecolors.ShadowBackgroundColor = new CellData<int>(fr[bggord], (int) rr[bggord]);
            shapecolors.LineColor = new CellData<int>(fr[lcord], (int) rr[lcord]);
            shapecolors.CharacterColor = new CellData<int>(fr[ccord], (int) rr[ccord]);

            return shapecolors;
        }

        public class ShapeColors
        {
            public VA.ShapeSheet.CellData<int> FillForegroundColor;
            public VA.ShapeSheet.CellData<int> FillBackgroundColor;
            public VA.ShapeSheet.CellData<int> ShadowForegroundColor;
            public VA.ShapeSheet.CellData<int> ShadowBackgroundColor;
            public VA.ShapeSheet.CellData<int> LineColor;
            public VA.ShapeSheet.CellData<int> CharacterColor;

            public override string ToString()
            {
                string s = typeof (ShapeColors).Name;
                return s;
            }
        }
    }
}