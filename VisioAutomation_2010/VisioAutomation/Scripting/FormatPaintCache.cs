using System.Collections.Generic;
using System.Linq;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting
{
    public class FormatPaintCache
    {
        public List<FormatPaintCell> Cells { get; }

        public FormatPaintCache()
        {
            this.Cells = new List<FormatPaintCell>();

            this.Add(SRCCON.FillBkgnd, FormatCategory.Fill, "FillBkgnd");
            this.Add(SRCCON.FillBkgndTrans, FormatCategory.Fill, "FillBkgndTrans");
            this.Add(SRCCON.FillForegnd, FormatCategory.Fill, "FillForegnd");
            this.Add(SRCCON.FillForegndTrans, FormatCategory.Fill, "FillForegndTrans");
            this.Add(SRCCON.FillPattern, FormatCategory.Fill, "FillPattern");

            this.Add(SRCCON.ShapeShdwObliqueAngle, FormatCategory.Shadow, "ShapeShdwObliqueAngle");
            this.Add(SRCCON.ShapeShdwOffsetX, FormatCategory.Shadow, "ShapeShdwOffsetX");
            this.Add(SRCCON.ShapeShdwOffsetY, FormatCategory.Shadow, "ShapeShdwOffsetY");
            this.Add(SRCCON.ShapeShdwScaleFactor, FormatCategory.Shadow, "ShapeShdwScaleFactor");
            this.Add(SRCCON.ShapeShdwType, FormatCategory.Shadow, "ShapeShdwType");
            this.Add(SRCCON.ShdwBkgnd, FormatCategory.Shadow, "ShdwBkgnd");
            this.Add(SRCCON.ShdwBkgndTrans, FormatCategory.Shadow, "ShdwBkgndTrans");
            this.Add(SRCCON.ShdwForegnd, FormatCategory.Shadow, "ShdwForegnd");
            this.Add(SRCCON.ShdwForegndTrans, FormatCategory.Shadow, "ShdwForegndTrans");
            this.Add(SRCCON.ShdwPattern, FormatCategory.Shadow, "ShdwPattern");

            this.Add(SRCCON.BeginArrow, FormatCategory.Line, "BeginArrow");
            this.Add(SRCCON.BeginArrowSize, FormatCategory.Line, "BeginArrowSize");
            this.Add(SRCCON.EndArrow, FormatCategory.Line, "EndArrow");
            this.Add(SRCCON.EndArrowSize, FormatCategory.Line, "EndArrowSize");
            this.Add(SRCCON.LineCap, FormatCategory.Line, "LineCap");
            this.Add(SRCCON.LineColor, FormatCategory.Line, "LineColor");
            this.Add(SRCCON.LineColorTrans, FormatCategory.Line, "LineColorTrans");
            this.Add(SRCCON.LinePattern, FormatCategory.Line, "LinePattern");
            this.Add(SRCCON.LineWeight, FormatCategory.Line, "LineWeight");
            this.Add(SRCCON.Rounding, FormatCategory.Line, "Rounding");

            this.Add(SRCCON.CharSize, FormatCategory.Character, "Char_Size");
            this.Add(SRCCON.CharLetterspace, FormatCategory.Character, "Char_Letterspace");
            this.Add(SRCCON.CharFontScale, FormatCategory.Character, "Char_FontScale");
            this.Add(SRCCON.CharStrikethru, FormatCategory.Character, "Char_Strikethru");
            this.Add(SRCCON.CharStrikethru, FormatCategory.Character, "Char_Strikethru");
            this.Add(SRCCON.CharFont, FormatCategory.Character, "Char_Font");
            this.Add(SRCCON.CharColorTrans, FormatCategory.Character, "Char_ColorTrans");
            this.Add(SRCCON.CharUseVertical, FormatCategory.Character, "Char_UseVertical");
            this.Add(SRCCON.CharCase, FormatCategory.Character, "Char_Case");
            this.Add(SRCCON.CharColor, FormatCategory.Character, "Char_Color");
            this.Add(SRCCON.CharComplexScriptFont, FormatCategory.Character, "Char_ComplexScriptFont");
            this.Add(SRCCON.CharComplexScriptSize, FormatCategory.Character, "Char_ComplexScriptSize");
            this.Add(SRCCON.CharRTLText, FormatCategory.Character, "Char_RTLText");
            this.Add(SRCCON.CharPerpendicular, FormatCategory.Character, "Char_Perpendicular");
            this.Add(SRCCON.CharOverline, FormatCategory.Character, "Char_Overline");
            this.Add(SRCCON.CharDoubleStrikethrough, FormatCategory.Character, "Char_DoubleStrikethrough");
            this.Add(SRCCON.CharDblUnderline, FormatCategory.Character, "Char_DblUnderline");
            this.Add(SRCCON.CharLangID, FormatCategory.Character, "Char_LangID");
            this.Add(SRCCON.CharLocale, FormatCategory.Character, "Char_Locale");
            this.Add(SRCCON.CharLocalizeFont, FormatCategory.Character, "Char_LocalizeFont");

            this.Add(SRCCON.Para_Bullet, FormatCategory.Paragraph, "Para_Bullet");
            this.Add(SRCCON.Para_BulletFont, FormatCategory.Paragraph, "Para_BulletFont");
            this.Add(SRCCON.Para_BulletFontSize, FormatCategory.Paragraph, "Para_BulletFontSize");
            this.Add(SRCCON.Para_BulletStr, FormatCategory.Paragraph, "Para_BulletStr");
            this.Add(SRCCON.Para_Flags, FormatCategory.Paragraph, "Para_Flags");
            this.Add(SRCCON.Para_HorzAlign, FormatCategory.Paragraph, "Para_HorzAlign");
            this.Add(SRCCON.Para_IndFirst, FormatCategory.Paragraph, "Para_IndFirst");
            this.Add(SRCCON.Para_IndLeft, FormatCategory.Paragraph, "Para_IndLeft");
            this.Add(SRCCON.Para_IndRight, FormatCategory.Paragraph, "Para_IndRight");
            this.Add(SRCCON.Para_LocalizeBulletFont, FormatCategory.Paragraph, "Para_LocalizeBulletFont");
            this.Add(SRCCON.Para_SpAfter, FormatCategory.Paragraph, "Para_SpAfter");
            this.Add(SRCCON.Para_SpBefore, FormatCategory.Paragraph, "Para_SpBefore");
            this.Add(SRCCON.Para_SpLine, FormatCategory.Paragraph, "Para_SpLine");
            this.Add(SRCCON.Para_TextPosAfterBullet, FormatCategory.Paragraph, "Para_TextPosAfterBullet");
        }

        public void Add(ShapeSheet.SRC src, FormatCategory category, string name)
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
            var query = new ShapeSheet.Query.CellQuery();
            var desired_cells = this.Cells.Where(cell => cell.MatchesCategory(category)).ToList();

            foreach (var cell in desired_cells)
            {
                query.AddCell(cell.SRC, null);
            }

            // Retrieve the values for the cells
            var dataset = query.GetCellData<string>(shape);

            // Now store the values
            for (int col = 0; col < query.CellColumns.Count; col++)
            {
                var cellrec = desired_cells[col];

                var result = dataset[col].Result;
                var formula = dataset[col].Formula;

                cellrec.Result = result;
                cellrec.Formula = formula.Value;
            }
        }

        public void PasteFormat(IVisio.Page page, IList<int> shapeids, FormatCategory category, bool applyformulas)
        {
            var update = new ShapeSheet.Update();

            foreach (var shape_id in shapeids)
            {
                foreach (var cellrec in this.Cells)
                {
                    if (!cellrec.MatchesCategory(category))
                    {
                        continue;
                    }

                    var sidsrc = new ShapeSheet.SIDSRC((short)shape_id, cellrec.SRC);

                    if (applyformulas)
                    {
                        update.SetFormula(sidsrc, cellrec.Formula);
                        
                    }
                    else
                    {
                        if (cellrec.Result != null)
                        {
                            update.SetFormula(sidsrc, cellrec.Result);
                        }
                    }
                }
            }

            update.Execute(page);
        }

        public FormatCategory GetAllFormatPaintFlags()
        {
            return FormatCategory.Fill | FormatCategory.Line | FormatCategory.Shadow | FormatCategory.Character | FormatCategory.Paragraph;
        }
    }
}