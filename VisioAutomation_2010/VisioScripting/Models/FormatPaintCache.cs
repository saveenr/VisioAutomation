using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
{
    public class FormatPaintCache
    {
        public List<FormatPaintCell> Cells { get; }

        public FormatPaintCache()
        {
            this.Cells = new List<FormatPaintCell>();

            this.Add(SrcConstants.FillBackground, FormatPaintCategory.Fill, "FillBkgnd");
            this.Add(SrcConstants.FillBackgroundTransparency, FormatPaintCategory.Fill, "FillBkgndTrans");
            this.Add(SrcConstants.FillForeground, FormatPaintCategory.Fill, "FillForegnd");
            this.Add(SrcConstants.FillForegroundTransparency, FormatPaintCategory.Fill, "FillForegndTrans");
            this.Add(SrcConstants.FillPattern, FormatPaintCategory.Fill, "FillPattern");

            this.Add(SrcConstants.FillShadowObliqueAngle, FormatPaintCategory.Shadow, "ShapeShdwObliqueAngle");
            this.Add(SrcConstants.FillShadowOffsetX, FormatPaintCategory.Shadow, "ShapeShdwOffsetX");
            this.Add(SrcConstants.FillShadowOffsetY, FormatPaintCategory.Shadow, "ShapeShdwOffsetY");
            this.Add(SrcConstants.FillShadowScaleFactor, FormatPaintCategory.Shadow, "ShapeShdwScaleFactor");
            this.Add(SrcConstants.FillShadowType, FormatPaintCategory.Shadow, "ShapeShdwType");
            this.Add(SrcConstants.FillShadowBackground, FormatPaintCategory.Shadow, "ShdwBkgnd");
            this.Add(SrcConstants.FillShadowBackgroundTransparency, FormatPaintCategory.Shadow, "ShdwBkgndTrans");
            this.Add(SrcConstants.FillShadowForeground, FormatPaintCategory.Shadow, "ShdwForegnd");
            this.Add(SrcConstants.FillShadowForegroundTransparency, FormatPaintCategory.Shadow, "ShdwForegndTrans");
            this.Add(SrcConstants.FillShadowPattern, FormatPaintCategory.Shadow, "ShdwPattern");

            this.Add(SrcConstants.LineBeginArrow, FormatPaintCategory.Line, "BeginArrow");
            this.Add(SrcConstants.LineBeginArrowSize, FormatPaintCategory.Line, "BeginArrowSize");
            this.Add(SrcConstants.LineEndArrow, FormatPaintCategory.Line, "EndArrow");
            this.Add(SrcConstants.LineEndArrowSize, FormatPaintCategory.Line, "EndArrowSize");
            this.Add(SrcConstants.LineCap, FormatPaintCategory.Line, "LineCap");
            this.Add(SrcConstants.LineColor, FormatPaintCategory.Line, "LineColor");
            this.Add(SrcConstants.LineColorTransparency, FormatPaintCategory.Line, "LineColorTrans");
            this.Add(SrcConstants.LinePattern, FormatPaintCategory.Line, "LinePattern");
            this.Add(SrcConstants.LineWeight, FormatPaintCategory.Line, "LineWeight");
            this.Add(SrcConstants.LineRounding, FormatPaintCategory.Line, "Rounding");

            this.Add(SrcConstants.CharSize, FormatPaintCategory.Character, "Char_Size");
            this.Add(SrcConstants.CharLetterspace, FormatPaintCategory.Character, "Char_Letterspace");
            this.Add(SrcConstants.CharFontScale, FormatPaintCategory.Character, "Char_FontScale");
            this.Add(SrcConstants.CharStrikethru, FormatPaintCategory.Character, "Char_Strikethru");
            this.Add(SrcConstants.CharStrikethru, FormatPaintCategory.Character, "Char_Strikethru");
            this.Add(SrcConstants.CharFont, FormatPaintCategory.Character, "Char_Font");
            this.Add(SrcConstants.CharColorTransparency, FormatPaintCategory.Character, "Char_ColorTrans");
            this.Add(SrcConstants.CharUseVertical, FormatPaintCategory.Character, "Char_UseVertical");
            this.Add(SrcConstants.CharCase, FormatPaintCategory.Character, "Char_Case");
            this.Add(SrcConstants.CharColor, FormatPaintCategory.Character, "Char_Color");
            this.Add(SrcConstants.CharComplexScriptFont, FormatPaintCategory.Character, "Char_ComplexScriptFont");
            this.Add(SrcConstants.CharComplexScriptSize, FormatPaintCategory.Character, "Char_ComplexScriptSize");
            this.Add(SrcConstants.CharRTLText, FormatPaintCategory.Character, "Char_RTLText");
            this.Add(SrcConstants.CharPerpendicular, FormatPaintCategory.Character, "Char_Perpendicular");
            this.Add(SrcConstants.CharOverline, FormatPaintCategory.Character, "Char_Overline");
            this.Add(SrcConstants.CharDoubleStrikethrough, FormatPaintCategory.Character, "Char_DoubleStrikethrough");
            this.Add(SrcConstants.CharDoubleUnderline, FormatPaintCategory.Character, "Char_DblUnderline");
            this.Add(SrcConstants.CharLangID, FormatPaintCategory.Character, "Char_LangID");
            this.Add(SrcConstants.CharLocale, FormatPaintCategory.Character, "Char_Locale");
            this.Add(SrcConstants.CharLocalizeFont, FormatPaintCategory.Character, "Char_LocalizeFont");

            this.Add(SrcConstants.ParaBullet, FormatPaintCategory.Paragraph, "Para_Bullet");
            this.Add(SrcConstants.ParaBulletFont, FormatPaintCategory.Paragraph, "Para_BulletFont");
            this.Add(SrcConstants.ParaBulletFontSize, FormatPaintCategory.Paragraph, "Para_BulletFontSize");
            this.Add(SrcConstants.ParaBulletString, FormatPaintCategory.Paragraph, "Para_BulletStr");
            this.Add(SrcConstants.ParaFlags, FormatPaintCategory.Paragraph, "Para_Flags");
            this.Add(SrcConstants.ParaHorizontalAlign, FormatPaintCategory.Paragraph, "Para_HorzAlign");
            this.Add(SrcConstants.ParaIndentFirst, FormatPaintCategory.Paragraph, "Para_IndFirst");
            this.Add(SrcConstants.ParaIndentLeft, FormatPaintCategory.Paragraph, "Para_IndLeft");
            this.Add(SrcConstants.ParaIndentRight, FormatPaintCategory.Paragraph, "Para_IndRight");
            this.Add(SrcConstants.ParaLocalizeBulletFont, FormatPaintCategory.Paragraph, "Para_LocalizeBulletFont");
            this.Add(SrcConstants.ParaSpacingAfter, FormatPaintCategory.Paragraph, "Para_SpAfter");
            this.Add(SrcConstants.ParaSpacingBefore, FormatPaintCategory.Paragraph, "Para_SpBefore");
            this.Add(SrcConstants.ParaSpacingLine, FormatPaintCategory.Paragraph, "Para_SpLine");
            this.Add(SrcConstants.ParaTextPosAfterBullet, FormatPaintCategory.Paragraph, "Para_TextPosAfterBullet");
        }

        public void Add(VisioAutomation.ShapeSheet.Src src, FormatPaintCategory paint_category, string name)
        {
            var format_cell = new FormatPaintCell(src, name, paint_category);
            this.Cells.Add(format_cell);
        }

        public void Clear()
        {
            foreach (var cell in this.Cells)
            {
                cell.Clear();
            }
        }
        
        public void CopyFormat(IVisio.Shape shape, FormatPaintCategory paint_category)
        {
            // Build the Query
            var query = new ShapeSheetQuery();
            var desired_cells = this.Cells.Where(cell => cell.MatchesCategory(paint_category)).ToList();

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
                cellrec.Formula = formula;
            }
        }

        public void PasteFormat(IVisio.Page page, IList<int> shapeids, FormatPaintCategory paint_category, bool applyformulas)
        {

            // Find all the cells that are going to be pasted
            var matching_cells = this.Cells.Where(c => c.MatchesCategory(paint_category)).ToArray();

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

        public FormatPaintCategory GetAllFormatPaintFlags()
        {
            return FormatPaintCategory.Fill | FormatPaintCategory.Line | FormatPaintCategory.Shadow | FormatPaintCategory.Character | FormatPaintCategory.Paragraph;
        }
    }
}