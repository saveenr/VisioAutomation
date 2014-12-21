using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting
{
    public class FormatPaintCache
    {
        public List<FormatPaintCell> Cells { get; private set; }

        public FormatPaintCache()
        {
            this.Cells = new List<FormatPaintCell>();

            this.Add(FormatCategory.Fill, "FillBkgnd", VA.ShapeSheet.SRCConstants.FillBkgnd);
            this.Add(FormatCategory.Fill, "FillBkgndTrans", VA.ShapeSheet.SRCConstants.FillBkgndTrans);
            this.Add(FormatCategory.Fill, "FillForegnd", VA.ShapeSheet.SRCConstants.FillForegnd);
            this.Add(FormatCategory.Fill, "FillForegndTrans", VA.ShapeSheet.SRCConstants.FillForegndTrans);
            this.Add(FormatCategory.Fill, "FillPattern", VA.ShapeSheet.SRCConstants.FillPattern);

            this.Add(FormatCategory.Shadow, "ShapeShdwObliqueAngle", VA.ShapeSheet.SRCConstants.ShapeShdwObliqueAngle);
            this.Add(FormatCategory.Shadow, "ShapeShdwOffsetX", VA.ShapeSheet.SRCConstants.ShapeShdwOffsetX);
            this.Add(FormatCategory.Shadow, "ShapeShdwOffsetY", VA.ShapeSheet.SRCConstants.ShapeShdwOffsetY);
            this.Add(FormatCategory.Shadow, "ShapeShdwScaleFactor", VA.ShapeSheet.SRCConstants.ShapeShdwScaleFactor);
            this.Add(FormatCategory.Shadow, "ShapeShdwType", VA.ShapeSheet.SRCConstants.ShapeShdwType);
            this.Add(FormatCategory.Shadow, "ShdwBkgnd", VA.ShapeSheet.SRCConstants.ShdwBkgnd);
            this.Add(FormatCategory.Shadow, "ShdwBkgndTrans", VA.ShapeSheet.SRCConstants.ShdwBkgndTrans);
            this.Add(FormatCategory.Shadow, "ShdwForegnd", VA.ShapeSheet.SRCConstants.ShdwForegnd);
            this.Add(FormatCategory.Shadow, "ShdwForegndTrans", VA.ShapeSheet.SRCConstants.ShdwForegndTrans);
            this.Add(FormatCategory.Shadow, "ShdwPattern", VA.ShapeSheet.SRCConstants.ShdwPattern);

            this.Add(FormatCategory.Line, "BeginArrow", VA.ShapeSheet.SRCConstants.BeginArrow);
            this.Add(FormatCategory.Line, "BeginArrowSize", VA.ShapeSheet.SRCConstants.BeginArrowSize);
            this.Add(FormatCategory.Line, "EndArrow", VA.ShapeSheet.SRCConstants.EndArrow);
            this.Add(FormatCategory.Line, "EndArrowSize", VA.ShapeSheet.SRCConstants.EndArrowSize);
            this.Add(FormatCategory.Line, "LineCap", VA.ShapeSheet.SRCConstants.LineCap);
            this.Add(FormatCategory.Line, "LineColor", VA.ShapeSheet.SRCConstants.LineColor);
            this.Add(FormatCategory.Line, "LineColorTrans", VA.ShapeSheet.SRCConstants.LineColorTrans);
            this.Add(FormatCategory.Line, "LinePattern", VA.ShapeSheet.SRCConstants.LinePattern);
            this.Add(FormatCategory.Line, "LineWeight", VA.ShapeSheet.SRCConstants.LineWeight);
            this.Add(FormatCategory.Line, "Rounding", VA.ShapeSheet.SRCConstants.Rounding);

            this.Add(FormatCategory.Character, "Char_Size", VA.ShapeSheet.SRCConstants.CharSize);
            this.Add(FormatCategory.Character, "Char_Letterspace", VA.ShapeSheet.SRCConstants.CharLetterspace);
            this.Add(FormatCategory.Character, "Char_FontScale", VA.ShapeSheet.SRCConstants.CharFontScale);
            this.Add(FormatCategory.Character, "Char_Strikethru", VA.ShapeSheet.SRCConstants.CharStrikethru);
            this.Add(FormatCategory.Character, "Char_Strikethru", VA.ShapeSheet.SRCConstants.CharStrikethru);
            this.Add(FormatCategory.Character, "Char_Font", VA.ShapeSheet.SRCConstants.CharFont);
            this.Add(FormatCategory.Character, "Char_ColorTrans", VA.ShapeSheet.SRCConstants.CharColorTrans);
            this.Add(FormatCategory.Character, "Char_UseVertical", VA.ShapeSheet.SRCConstants.CharUseVertical);
            this.Add(FormatCategory.Character, "Char_Case", VA.ShapeSheet.SRCConstants.CharCase);
            this.Add(FormatCategory.Character, "Char_Color", VA.ShapeSheet.SRCConstants.CharColor);
            this.Add(FormatCategory.Character, "Char_ComplexScriptFont", VA.ShapeSheet.SRCConstants.CharComplexScriptFont);
            this.Add(FormatCategory.Character, "Char_ComplexScriptSize", VA.ShapeSheet.SRCConstants.CharComplexScriptSize);
            this.Add(FormatCategory.Character, "Char_RTLText", VA.ShapeSheet.SRCConstants.CharRTLText);
            this.Add(FormatCategory.Character, "Char_Perpendicular", VA.ShapeSheet.SRCConstants.CharPerpendicular);
            this.Add(FormatCategory.Character, "Char_Overline", VA.ShapeSheet.SRCConstants.CharOverline);
            this.Add(FormatCategory.Character, "Char_DoubleStrikethrough", VA.ShapeSheet.SRCConstants.CharDoubleStrikethrough);
            this.Add(FormatCategory.Character, "Char_DblUnderline", VA.ShapeSheet.SRCConstants.CharDblUnderline);
            this.Add(FormatCategory.Character, "Char_LangID", VA.ShapeSheet.SRCConstants.CharLangID);
            this.Add(FormatCategory.Character, "Char_Locale", VA.ShapeSheet.SRCConstants.CharLocale);
            this.Add(FormatCategory.Character, "Char_LocalizeFont", VA.ShapeSheet.SRCConstants.CharLocalizeFont);

            this.Add(FormatCategory.Paragraph, "Para_Bullet", VA.ShapeSheet.SRCConstants.Para_Bullet);
            this.Add(FormatCategory.Paragraph, "Para_BulletFont", VA.ShapeSheet.SRCConstants.Para_BulletFont);
            this.Add(FormatCategory.Paragraph, "Para_BulletFontSize", VA.ShapeSheet.SRCConstants.Para_BulletFontSize);
            this.Add(FormatCategory.Paragraph, "Para_BulletStr", VA.ShapeSheet.SRCConstants.Para_BulletStr);
            this.Add(FormatCategory.Paragraph, "Para_Flags", VA.ShapeSheet.SRCConstants.Para_Flags);
            this.Add(FormatCategory.Paragraph, "Para_HorzAlign", VA.ShapeSheet.SRCConstants.Para_HorzAlign);
            this.Add(FormatCategory.Paragraph, "Para_IndFirst", VA.ShapeSheet.SRCConstants.Para_IndFirst);
            this.Add(FormatCategory.Paragraph, "Para_IndLeft", VA.ShapeSheet.SRCConstants.Para_IndLeft);
            this.Add(FormatCategory.Paragraph, "Para_IndRight", VA.ShapeSheet.SRCConstants.Para_IndRight);
            this.Add(FormatCategory.Paragraph, "Para_LocalizeBulletFont", VA.ShapeSheet.SRCConstants.Para_LocalizeBulletFont);
            this.Add(FormatCategory.Paragraph, "Para_SpAfter", VA.ShapeSheet.SRCConstants.Para_SpAfter);
            this.Add(FormatCategory.Paragraph, "Para_SpBefore", VA.ShapeSheet.SRCConstants.Para_SpBefore);
            this.Add(FormatCategory.Paragraph, "Para_SpLine", VA.ShapeSheet.SRCConstants.Para_SpLine);
            this.Add(FormatCategory.Paragraph, "Para_TextPosAfterBullet", VA.ShapeSheet.SRCConstants.Para_TextPosAfterBullet);
        }

        public void Add(FormatCategory category, string name, VA.ShapeSheet.SRC src)
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
            var query = new VA.ShapeSheet.Query.CellQuery();
            var desired_cells = this.Cells.Where(cell => cell.MatchesCategory(category)).ToList();

            foreach (var cell in desired_cells)
            {
                query.Columns.Add(cell.SRC,null);
            }

            // Retrieve the values for the cells
            var dataset = query.GetFormulasAndResults<string>(shape);

            // Now store the values
            for (int col = 0; col < query.Columns.Count; col++)
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
            var update = new VA.ShapeSheet.Update();

            foreach (var shape_id in shapeids)
            {
                foreach (var cellrec in this.Cells)
                {
                    if (!cellrec.MatchesCategory(category))
                    {
                        continue;
                    }

                    var sidsrc = new VA.ShapeSheet.SIDSRC((short)shape_id, cellrec.SRC);

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