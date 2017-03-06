using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.FormatPaint
{
    public class FormatPaintCache
    {
        public List<FormatPaintCell> Cells { get; }

        public FormatPaintCache()
        {
            this.Cells = new List<FormatPaintCell>();

            this.Add(SrcConstants.FillBackground, FormatCategory.Fill, "FillBkgnd");
            this.Add(SrcConstants.FillBackgroundTransparency, FormatCategory.Fill, "FillBkgndTrans");
            this.Add(SrcConstants.FillForeground, FormatCategory.Fill, "FillForegnd");
            this.Add(SrcConstants.FillForegroundTransparency, FormatCategory.Fill, "FillForegndTrans");
            this.Add(SrcConstants.FillPattern, FormatCategory.Fill, "FillPattern");

            this.Add(SrcConstants.FillShadowObliqueAngle, FormatCategory.Shadow, "ShapeShdwObliqueAngle");
            this.Add(SrcConstants.FillShadowOffsetX, FormatCategory.Shadow, "ShapeShdwOffsetX");
            this.Add(SrcConstants.FillShadowOffsetY, FormatCategory.Shadow, "ShapeShdwOffsetY");
            this.Add(SrcConstants.FillShadowScaleFactor, FormatCategory.Shadow, "ShapeShdwScaleFactor");
            this.Add(SrcConstants.FillShadowType, FormatCategory.Shadow, "ShapeShdwType");
            this.Add(SrcConstants.FillShadowBackground, FormatCategory.Shadow, "ShdwBkgnd");
            this.Add(SrcConstants.FillShadowBackgroundTransparency, FormatCategory.Shadow, "ShdwBkgndTrans");
            this.Add(SrcConstants.FillShadowForeground, FormatCategory.Shadow, "ShdwForegnd");
            this.Add(SrcConstants.FillShadowForegroundTransparency, FormatCategory.Shadow, "ShdwForegndTrans");
            this.Add(SrcConstants.FillShadowPattern, FormatCategory.Shadow, "ShdwPattern");

            this.Add(SrcConstants.LineBeginArrow, FormatCategory.Line, "BeginArrow");
            this.Add(SrcConstants.LineBeginArrowSize, FormatCategory.Line, "BeginArrowSize");
            this.Add(SrcConstants.LineEndArrow, FormatCategory.Line, "EndArrow");
            this.Add(SrcConstants.LineEndArrowSize, FormatCategory.Line, "EndArrowSize");
            this.Add(SrcConstants.LineCap, FormatCategory.Line, "LineCap");
            this.Add(SrcConstants.LineColor, FormatCategory.Line, "LineColor");
            this.Add(SrcConstants.LineColorTransparency, FormatCategory.Line, "LineColorTrans");
            this.Add(SrcConstants.LinePattern, FormatCategory.Line, "LinePattern");
            this.Add(SrcConstants.LineWeight, FormatCategory.Line, "LineWeight");
            this.Add(SrcConstants.LineRounding, FormatCategory.Line, "Rounding");

            this.Add(SrcConstants.CharSize, FormatCategory.Character, "Char_Size");
            this.Add(SrcConstants.CharLetterspace, FormatCategory.Character, "Char_Letterspace");
            this.Add(SrcConstants.CharFontScale, FormatCategory.Character, "Char_FontScale");
            this.Add(SrcConstants.CharStrikethru, FormatCategory.Character, "Char_Strikethru");
            this.Add(SrcConstants.CharStrikethru, FormatCategory.Character, "Char_Strikethru");
            this.Add(SrcConstants.CharFont, FormatCategory.Character, "Char_Font");
            this.Add(SrcConstants.CharColorTransparency, FormatCategory.Character, "Char_ColorTrans");
            this.Add(SrcConstants.CharUseVertical, FormatCategory.Character, "Char_UseVertical");
            this.Add(SrcConstants.CharCase, FormatCategory.Character, "Char_Case");
            this.Add(SrcConstants.CharColor, FormatCategory.Character, "Char_Color");
            this.Add(SrcConstants.CharComplexScriptFont, FormatCategory.Character, "Char_ComplexScriptFont");
            this.Add(SrcConstants.CharComplexScriptSize, FormatCategory.Character, "Char_ComplexScriptSize");
            this.Add(SrcConstants.CharRTLText, FormatCategory.Character, "Char_RTLText");
            this.Add(SrcConstants.CharPerpendicular, FormatCategory.Character, "Char_Perpendicular");
            this.Add(SrcConstants.CharOverline, FormatCategory.Character, "Char_Overline");
            this.Add(SrcConstants.CharDoubleStrikethrough, FormatCategory.Character, "Char_DoubleStrikethrough");
            this.Add(SrcConstants.CharDoubleUnderline, FormatCategory.Character, "Char_DblUnderline");
            this.Add(SrcConstants.CharLangID, FormatCategory.Character, "Char_LangID");
            this.Add(SrcConstants.CharLocale, FormatCategory.Character, "Char_Locale");
            this.Add(SrcConstants.CharLocalizeFont, FormatCategory.Character, "Char_LocalizeFont");

            this.Add(SrcConstants.ParaBullet, FormatCategory.Paragraph, "Para_Bullet");
            this.Add(SrcConstants.ParaBulletFont, FormatCategory.Paragraph, "Para_BulletFont");
            this.Add(SrcConstants.ParaBulletFontSize, FormatCategory.Paragraph, "Para_BulletFontSize");
            this.Add(SrcConstants.ParaBulletStr, FormatCategory.Paragraph, "Para_BulletStr");
            this.Add(SrcConstants.ParaFlags, FormatCategory.Paragraph, "Para_Flags");
            this.Add(SrcConstants.ParaHorizontalAlign, FormatCategory.Paragraph, "Para_HorzAlign");
            this.Add(SrcConstants.ParaIndentFirst, FormatCategory.Paragraph, "Para_IndFirst");
            this.Add(SrcConstants.ParaIndentLeft, FormatCategory.Paragraph, "Para_IndLeft");
            this.Add(SrcConstants.ParaIndentRight, FormatCategory.Paragraph, "Para_IndRight");
            this.Add(SrcConstants.ParaLocalizeBulletFont, FormatCategory.Paragraph, "Para_LocalizeBulletFont");
            this.Add(SrcConstants.ParaSpacingAfter, FormatCategory.Paragraph, "Para_SpAfter");
            this.Add(SrcConstants.ParaSpacingBefore, FormatCategory.Paragraph, "Para_SpBefore");
            this.Add(SrcConstants.ParaSpacingLine, FormatCategory.Paragraph, "Para_SpLine");
            this.Add(SrcConstants.ParaTextPosAfterBullet, FormatCategory.Paragraph, "Para_TextPosAfterBullet");
        }

        public void Add(VisioAutomation.ShapeSheet.Src src, FormatCategory category, string name)
        {
            var format_cell = new FormatPaintCell(src, name, category);
            this.Cells.Add(format_cell);
        }

        public void Clear()
        {
            foreach (var cell in this.Cells)
            {
                cell.Clear();
            }
        }
        
        public void CopyFormat(IVisio.Shape shape, FormatCategory category)
        {
            // Build the Query
            var query = new ShapeSheetQuery();
            var desired_cells = this.Cells.Where(cell => cell.MatchesCategory(category)).ToList();

            foreach (var cell in desired_cells)
            {
                query.AddCell(cell.Src, null);
            }

            // Retrieve the values for the cells
            var dataset = query.GetFormulasAndResults(shape);

            // Now store the values
            for (int col = 0; col < query.Cells.Count; col++)
            {
                var result = dataset.Cells[col].Result;
                var formula = dataset.Cells[col].Formula;

                var cellrec = desired_cells[col];
                cellrec.Result = result;
                cellrec.Formula = formula.Value;
            }
        }

        public void PasteFormat(IVisio.Page page, IList<int> shapeids, FormatCategory category, bool applyformulas)
        {

            // Find all the cells that are going to be pasted
            var matching_cells = this.Cells.Where(c => c.MatchesCategory(category)).ToArray();

            // Apply those matched cells to each shape
            var writer = new SidSrcWriter();
            foreach (var shape_id in shapeids)
            {
                foreach (var cell in matching_cells)
                {
                    var sidsrc = new VisioAutomation.ShapeSheet.SidSrc((short) shape_id, cell.Src);
                    var new_formula = applyformulas ? cell.Formula : cell.Result;
                    writer.SetFormula(sidsrc, new_formula);
                }
            }

            writer.Commit(page);
        }

        public FormatCategory GetAllFormatPaintFlags()
        {
            return FormatCategory.Fill | FormatCategory.Line | FormatCategory.Shadow | FormatCategory.Character | FormatCategory.Paragraph;
        }
    }
}