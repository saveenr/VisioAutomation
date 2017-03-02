using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.FormatPaint
{
    public class FormatPaintCache
    {
        public List<FormatPaintCell> Cells { get; }

        public FormatPaintCache()
        {
            this.Cells = new List<FormatPaintCell>();

            this.Add(SrcConstants.FillBkgnd, FormatCategory.Fill, "FillBkgnd");
            this.Add(SrcConstants.FillBkgndTrans, FormatCategory.Fill, "FillBkgndTrans");
            this.Add(SrcConstants.FillForegnd, FormatCategory.Fill, "FillForegnd");
            this.Add(SrcConstants.FillForegndTrans, FormatCategory.Fill, "FillForegndTrans");
            this.Add(SrcConstants.FillPattern, FormatCategory.Fill, "FillPattern");

            this.Add(SrcConstants.ShapeShdwObliqueAngle, FormatCategory.Shadow, "ShapeShdwObliqueAngle");
            this.Add(SrcConstants.ShapeShdwOffsetX, FormatCategory.Shadow, "ShapeShdwOffsetX");
            this.Add(SrcConstants.ShapeShdwOffsetY, FormatCategory.Shadow, "ShapeShdwOffsetY");
            this.Add(SrcConstants.ShapeShdwScaleFactor, FormatCategory.Shadow, "ShapeShdwScaleFactor");
            this.Add(SrcConstants.ShapeShdwType, FormatCategory.Shadow, "ShapeShdwType");
            this.Add(SrcConstants.ShdwBkgnd, FormatCategory.Shadow, "ShdwBkgnd");
            this.Add(SrcConstants.ShdwBkgndTrans, FormatCategory.Shadow, "ShdwBkgndTrans");
            this.Add(SrcConstants.ShdwForegnd, FormatCategory.Shadow, "ShdwForegnd");
            this.Add(SrcConstants.ShdwForegndTrans, FormatCategory.Shadow, "ShdwForegndTrans");
            this.Add(SrcConstants.ShdwPattern, FormatCategory.Shadow, "ShdwPattern");

            this.Add(SrcConstants.BeginArrow, FormatCategory.Line, "BeginArrow");
            this.Add(SrcConstants.BeginArrowSize, FormatCategory.Line, "BeginArrowSize");
            this.Add(SrcConstants.EndArrow, FormatCategory.Line, "EndArrow");
            this.Add(SrcConstants.EndArrowSize, FormatCategory.Line, "EndArrowSize");
            this.Add(SrcConstants.LineCap, FormatCategory.Line, "LineCap");
            this.Add(SrcConstants.LineColor, FormatCategory.Line, "LineColor");
            this.Add(SrcConstants.LineColorTrans, FormatCategory.Line, "LineColorTrans");
            this.Add(SrcConstants.LinePattern, FormatCategory.Line, "LinePattern");
            this.Add(SrcConstants.LineWeight, FormatCategory.Line, "LineWeight");
            this.Add(SrcConstants.Rounding, FormatCategory.Line, "Rounding");

            this.Add(SrcConstants.CharSize, FormatCategory.Character, "Char_Size");
            this.Add(SrcConstants.CharLetterspace, FormatCategory.Character, "Char_Letterspace");
            this.Add(SrcConstants.CharFontScale, FormatCategory.Character, "Char_FontScale");
            this.Add(SrcConstants.CharStrikethru, FormatCategory.Character, "Char_Strikethru");
            this.Add(SrcConstants.CharStrikethru, FormatCategory.Character, "Char_Strikethru");
            this.Add(SrcConstants.CharFont, FormatCategory.Character, "Char_Font");
            this.Add(SrcConstants.CharColorTrans, FormatCategory.Character, "Char_ColorTrans");
            this.Add(SrcConstants.CharUseVertical, FormatCategory.Character, "Char_UseVertical");
            this.Add(SrcConstants.CharCase, FormatCategory.Character, "Char_Case");
            this.Add(SrcConstants.CharColor, FormatCategory.Character, "Char_Color");
            this.Add(SrcConstants.CharComplexScriptFont, FormatCategory.Character, "Char_ComplexScriptFont");
            this.Add(SrcConstants.CharComplexScriptSize, FormatCategory.Character, "Char_ComplexScriptSize");
            this.Add(SrcConstants.CharRTLText, FormatCategory.Character, "Char_RTLText");
            this.Add(SrcConstants.CharPerpendicular, FormatCategory.Character, "Char_Perpendicular");
            this.Add(SrcConstants.CharOverline, FormatCategory.Character, "Char_Overline");
            this.Add(SrcConstants.CharDoubleStrikethrough, FormatCategory.Character, "Char_DoubleStrikethrough");
            this.Add(SrcConstants.CharDblUnderline, FormatCategory.Character, "Char_DblUnderline");
            this.Add(SrcConstants.CharLangID, FormatCategory.Character, "Char_LangID");
            this.Add(SrcConstants.CharLocale, FormatCategory.Character, "Char_Locale");
            this.Add(SrcConstants.CharLocalizeFont, FormatCategory.Character, "Char_LocalizeFont");

            this.Add(SrcConstants.Para_Bullet, FormatCategory.Paragraph, "Para_Bullet");
            this.Add(SrcConstants.Para_BulletFont, FormatCategory.Paragraph, "Para_BulletFont");
            this.Add(SrcConstants.Para_BulletFontSize, FormatCategory.Paragraph, "Para_BulletFontSize");
            this.Add(SrcConstants.Para_BulletStr, FormatCategory.Paragraph, "Para_BulletStr");
            this.Add(SrcConstants.Para_Flags, FormatCategory.Paragraph, "Para_Flags");
            this.Add(SrcConstants.Para_HorzAlign, FormatCategory.Paragraph, "Para_HorzAlign");
            this.Add(SrcConstants.Para_IndFirst, FormatCategory.Paragraph, "Para_IndFirst");
            this.Add(SrcConstants.Para_IndLeft, FormatCategory.Paragraph, "Para_IndLeft");
            this.Add(SrcConstants.Para_IndRight, FormatCategory.Paragraph, "Para_IndRight");
            this.Add(SrcConstants.Para_LocalizeBulletFont, FormatCategory.Paragraph, "Para_LocalizeBulletFont");
            this.Add(SrcConstants.Para_SpAfter, FormatCategory.Paragraph, "Para_SpAfter");
            this.Add(SrcConstants.Para_SpBefore, FormatCategory.Paragraph, "Para_SpBefore");
            this.Add(SrcConstants.Para_SpLine, FormatCategory.Paragraph, "Para_SpLine");
            this.Add(SrcConstants.Para_TextPosAfterBullet, FormatCategory.Paragraph, "Para_TextPosAfterBullet");
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
                query.AddCell(cell.SRC, null);
            }

            // Retrieve the values for the cells
            var surface = new ShapeSheetSurface(shape);
            var dataset = query.GetFormulasAndResults(surface);

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
            var writer = new ShapeSheetWriter();
            foreach (var shape_id in shapeids)
            {
                foreach (var cell in matching_cells)
                {
                    var sidsrc = new VisioAutomation.ShapeSheet.SidSrc((short) shape_id, cell.SRC);
                    var new_formula = applyformulas ? cell.Formula : cell.Result;
                    writer.SetFormula(sidsrc, new_formula);
                }
            }

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(page);
            writer.Commit(surface);
        }

        public FormatCategory GetAllFormatPaintFlags()
        {
            return FormatCategory.Fill | FormatCategory.Line | FormatCategory.Shadow | FormatCategory.Character | FormatCategory.Paragraph;
        }
    }
}