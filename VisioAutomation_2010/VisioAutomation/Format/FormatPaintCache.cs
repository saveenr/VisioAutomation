using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Format
{
    public class FormatPaintCache
    {
        public List<VA.Format.FormatPaintCell> Cells { get; private set; }

        public FormatPaintCache()
        {
            this.Cells = new List<VA.Format.FormatPaintCell>();

            this.Add(VA.Format.FormatCategory.Fill, "FillBkgnd", VA.ShapeSheet.SRCConstants.FillBkgnd);
            this.Add(VA.Format.FormatCategory.Fill, "FillBkgndTrans", VA.ShapeSheet.SRCConstants.FillBkgndTrans);
            this.Add(VA.Format.FormatCategory.Fill, "FillForegnd", VA.ShapeSheet.SRCConstants.FillForegnd);
            this.Add(VA.Format.FormatCategory.Fill, "FillForegndTrans", VA.ShapeSheet.SRCConstants.FillForegndTrans);
            this.Add(VA.Format.FormatCategory.Fill, "FillPattern", VA.ShapeSheet.SRCConstants.FillPattern);

            this.Add(VA.Format.FormatCategory.Shadow, "ShapeShdwObliqueAngle", VA.ShapeSheet.SRCConstants.ShapeShdwObliqueAngle);
            this.Add(VA.Format.FormatCategory.Shadow, "ShapeShdwOffsetX", VA.ShapeSheet.SRCConstants.ShapeShdwOffsetX);
            this.Add(VA.Format.FormatCategory.Shadow, "ShapeShdwOffsetY", VA.ShapeSheet.SRCConstants.ShapeShdwOffsetY);
            this.Add(VA.Format.FormatCategory.Shadow, "ShapeShdwScaleFactor", VA.ShapeSheet.SRCConstants.ShapeShdwScaleFactor);
            this.Add(VA.Format.FormatCategory.Shadow, "ShapeShdwType", VA.ShapeSheet.SRCConstants.ShapeShdwType);
            this.Add(VA.Format.FormatCategory.Shadow, "ShdwBkgnd", VA.ShapeSheet.SRCConstants.ShdwBkgnd);
            this.Add(VA.Format.FormatCategory.Shadow, "ShdwBkgndTrans", VA.ShapeSheet.SRCConstants.ShdwBkgndTrans);
            this.Add(VA.Format.FormatCategory.Shadow, "ShdwForegnd", VA.ShapeSheet.SRCConstants.ShdwForegnd);
            this.Add(VA.Format.FormatCategory.Shadow, "ShdwForegndTrans", VA.ShapeSheet.SRCConstants.ShdwForegndTrans);
            this.Add(VA.Format.FormatCategory.Shadow, "ShdwPattern", VA.ShapeSheet.SRCConstants.ShdwPattern);

            this.Add(VA.Format.FormatCategory.Line, "BeginArrow", VA.ShapeSheet.SRCConstants.BeginArrow);
            this.Add(VA.Format.FormatCategory.Line, "BeginArrowSize", VA.ShapeSheet.SRCConstants.BeginArrowSize);
            this.Add(VA.Format.FormatCategory.Line, "EndArrow", VA.ShapeSheet.SRCConstants.EndArrow);
            this.Add(VA.Format.FormatCategory.Line, "EndArrowSize", VA.ShapeSheet.SRCConstants.EndArrowSize);
            this.Add(VA.Format.FormatCategory.Line, "LineCap", VA.ShapeSheet.SRCConstants.LineCap);
            this.Add(VA.Format.FormatCategory.Line, "LineColor", VA.ShapeSheet.SRCConstants.LineColor);
            this.Add(VA.Format.FormatCategory.Line, "LineColorTrans", VA.ShapeSheet.SRCConstants.LineColorTrans);
            this.Add(VA.Format.FormatCategory.Line, "LinePattern", VA.ShapeSheet.SRCConstants.LinePattern);
            this.Add(VA.Format.FormatCategory.Line, "LineWeight", VA.ShapeSheet.SRCConstants.LineWeight);
            this.Add(VA.Format.FormatCategory.Line, "Rounding", VA.ShapeSheet.SRCConstants.Rounding);

            this.Add(VA.Format.FormatCategory.Character, "Char_Size", VA.ShapeSheet.SRCConstants.CharSize);
            this.Add(VA.Format.FormatCategory.Character, "Char_Letterspace", VA.ShapeSheet.SRCConstants.CharLetterspace);
            this.Add(VA.Format.FormatCategory.Character, "Char_FontScale", VA.ShapeSheet.SRCConstants.CharFontScale);
            this.Add(VA.Format.FormatCategory.Character, "Char_Strikethru", VA.ShapeSheet.SRCConstants.CharStrikethru);
            this.Add(VA.Format.FormatCategory.Character, "Char_Strikethru", VA.ShapeSheet.SRCConstants.CharStrikethru);
            this.Add(VA.Format.FormatCategory.Character, "Char_Font", VA.ShapeSheet.SRCConstants.CharFont);
            this.Add(VA.Format.FormatCategory.Character, "Char_ColorTrans", VA.ShapeSheet.SRCConstants.CharColorTrans);
            this.Add(VA.Format.FormatCategory.Character, "Char_UseVertical", VA.ShapeSheet.SRCConstants.CharUseVertical);
            this.Add(VA.Format.FormatCategory.Character, "Char_Case", VA.ShapeSheet.SRCConstants.CharCase);
            this.Add(VA.Format.FormatCategory.Character, "Char_Color", VA.ShapeSheet.SRCConstants.CharColor);
            this.Add(VA.Format.FormatCategory.Character, "Char_ComplexScriptFont", VA.ShapeSheet.SRCConstants.CharComplexScriptFont);
            this.Add(VA.Format.FormatCategory.Character, "Char_ComplexScriptSize", VA.ShapeSheet.SRCConstants.CharComplexScriptSize);
            this.Add(VA.Format.FormatCategory.Character, "Char_RTLText", VA.ShapeSheet.SRCConstants.CharRTLText);
            this.Add(VA.Format.FormatCategory.Character, "Char_Perpendicular", VA.ShapeSheet.SRCConstants.CharPerpendicular);
            this.Add(VA.Format.FormatCategory.Character, "Char_Overline", VA.ShapeSheet.SRCConstants.CharOverline);
            this.Add(VA.Format.FormatCategory.Character, "Char_DoubleStrikethrough", VA.ShapeSheet.SRCConstants.CharDoubleStrikethrough);
            this.Add(VA.Format.FormatCategory.Character, "Char_DblUnderline", VA.ShapeSheet.SRCConstants.CharDblUnderline);
            this.Add(VA.Format.FormatCategory.Character, "Char_LangID", VA.ShapeSheet.SRCConstants.CharLangID);
            this.Add(VA.Format.FormatCategory.Character, "Char_Locale", VA.ShapeSheet.SRCConstants.CharLocale);
            this.Add(VA.Format.FormatCategory.Character, "Char_LocalizeFont", VA.ShapeSheet.SRCConstants.CharLocalizeFont);

            this.Add(VA.Format.FormatCategory.Paragraph, "Para_Bullet", VA.ShapeSheet.SRCConstants.Para_Bullet);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_BulletFont", VA.ShapeSheet.SRCConstants.Para_BulletFont);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_BulletFontSize", VA.ShapeSheet.SRCConstants.Para_BulletFontSize);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_BulletStr", VA.ShapeSheet.SRCConstants.Para_BulletStr);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_Flags", VA.ShapeSheet.SRCConstants.Para_Flags);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_HorzAlign", VA.ShapeSheet.SRCConstants.Para_HorzAlign);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_IndFirst", VA.ShapeSheet.SRCConstants.Para_IndFirst);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_IndLeft", VA.ShapeSheet.SRCConstants.Para_IndLeft);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_IndRight", VA.ShapeSheet.SRCConstants.Para_IndRight);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_LocalizeBulletFont", VA.ShapeSheet.SRCConstants.Para_LocalizeBulletFont);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_SpAfter", VA.ShapeSheet.SRCConstants.Para_SpAfter);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_SpBefore", VA.ShapeSheet.SRCConstants.Para_SpBefore);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_SpLine", VA.ShapeSheet.SRCConstants.Para_SpLine);
            this.Add(VA.Format.FormatCategory.Paragraph, "Para_TextPosAfterBullet", VA.ShapeSheet.SRCConstants.Para_TextPosAfterBullet);
        }

        public void Add(VA.Format.FormatCategory category, string name, VA.ShapeSheet.SRC src)
        {
            var format_cell = new VA.Format.FormatPaintCell(src, name, category);
            this.Cells.Add(format_cell);
        }

        public void Clear()
        {
            foreach (var cell in this.Cells)
            {
                cell.Clear();
            }
        }
        
        public void CopyFormat(IVisio.Shape shape, VA.Format.FormatCategory category)
        {
            // Build the Query
            var query = new VA.ShapeSheet.Query.CellQuery();
            var desired_cells = this.Cells.Where(cell => cell.MatchesCategory(category)).ToList();

            foreach (var cell in desired_cells)
            {
                query.AddColumn(cell.SRC,null);
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

        public void PasteFormat(IVisio.Page page, IList<int> shapeids, VA.Format.FormatCategory category, bool applyformulas)
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