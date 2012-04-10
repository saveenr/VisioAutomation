using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
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

            this.Add(VA.Format.FormatCategory.Fill, VA.ShapeSheet.SRCConstants.FillBkgnd);
            this.Add(VA.Format.FormatCategory.Fill, VA.ShapeSheet.SRCConstants.FillBkgndTrans);
            this.Add(VA.Format.FormatCategory.Fill, VA.ShapeSheet.SRCConstants.FillForegnd);
            this.Add(VA.Format.FormatCategory.Fill, VA.ShapeSheet.SRCConstants.FillForegndTrans);
            this.Add(VA.Format.FormatCategory.Fill, VA.ShapeSheet.SRCConstants.FillPattern);

            this.Add(VA.Format.FormatCategory.Shadow, VA.ShapeSheet.SRCConstants.ShapeShdwObliqueAngle);
            this.Add(VA.Format.FormatCategory.Shadow, VA.ShapeSheet.SRCConstants.ShapeShdwOffsetX);
            this.Add(VA.Format.FormatCategory.Shadow, VA.ShapeSheet.SRCConstants.ShapeShdwOffsetY);
            this.Add(VA.Format.FormatCategory.Shadow, VA.ShapeSheet.SRCConstants.ShapeShdwScaleFactor);
            this.Add(VA.Format.FormatCategory.Shadow, VA.ShapeSheet.SRCConstants.ShapeShdwType);
            this.Add(VA.Format.FormatCategory.Shadow, VA.ShapeSheet.SRCConstants.ShdwBkgnd);
            this.Add(VA.Format.FormatCategory.Shadow, VA.ShapeSheet.SRCConstants.ShdwBkgndTrans);
            this.Add(VA.Format.FormatCategory.Shadow, VA.ShapeSheet.SRCConstants.ShdwForegnd);
            this.Add(VA.Format.FormatCategory.Shadow, VA.ShapeSheet.SRCConstants.ShdwForegndTrans);
            this.Add(VA.Format.FormatCategory.Shadow, VA.ShapeSheet.SRCConstants.ShdwPattern);

            this.Add(VA.Format.FormatCategory.Line, VA.ShapeSheet.SRCConstants.BeginArrow);
            this.Add(VA.Format.FormatCategory.Line, VA.ShapeSheet.SRCConstants.BeginArrowSize);
            this.Add(VA.Format.FormatCategory.Line, VA.ShapeSheet.SRCConstants.EndArrow);
            this.Add(VA.Format.FormatCategory.Line, VA.ShapeSheet.SRCConstants.EndArrowSize);
            this.Add(VA.Format.FormatCategory.Line, VA.ShapeSheet.SRCConstants.LineCap);
            this.Add(VA.Format.FormatCategory.Line, VA.ShapeSheet.SRCConstants.LineColor);
            this.Add(VA.Format.FormatCategory.Line, VA.ShapeSheet.SRCConstants.LineColorTrans);
            this.Add(VA.Format.FormatCategory.Line, VA.ShapeSheet.SRCConstants.LinePattern);
            this.Add(VA.Format.FormatCategory.Line, VA.ShapeSheet.SRCConstants.LineWeight);
            this.Add(VA.Format.FormatCategory.Line, VA.ShapeSheet.SRCConstants.Rounding);

            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_Size);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_Letterspace);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_FontScale);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_Strikethru);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_Style);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_Font);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_ColorTrans);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_UseVertical);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_Case);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_Color);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_ComplexScriptFont);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_ComplexScriptSize);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_RTLText);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_Perpendicular);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_Overline);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_DoubleStrikethrough);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_DblUnderline);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_LangID);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_Locale);
            this.Add(VA.Format.FormatCategory.Character, VA.ShapeSheet.SRCConstants.Char_LocalizeFont);

            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_Bullet);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_BulletFont);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_BulletFontSize);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_BulletStr);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_Flags);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_HorzAlign);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_IndFirst);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_IndLeft);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_IndRight);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_LocalizeBulletFont);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_SpAfter);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_SpBefore);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_SpLine);
            this.Add(VA.Format.FormatCategory.Paragraph, VA.ShapeSheet.SRCConstants.Para_TextPosAfterBullet);
        }

        public void Add(VA.Format.FormatCategory Category, VA.ShapeSheet.SRC src)
        {
            var format_cell = new VA.Format.FormatPaintCell(src, Category);
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
                query.AddColumn(cell.SRC);
            }

            // Retrieve the values for the cells
            var dataset = query.GetFormulasAndResults<double>(shape);

            // Now store the values
            for (int col = 0; col < query.Columns.Count; col++)
            {
                var cellrec = desired_cells[col];

                var result = dataset[0, col].Result;
                var formula = dataset[0, col].Formula;

                cellrec.Result = result;
                cellrec.Formula = formula.Value;
            }
        }

        public void PasteFormat(IVisio.Page page, IList<int> shapeids, VA.Format.FormatCategory category)
        {
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            foreach (var shape_id in shapeids)
            {
                foreach (var cellrec in this.Cells)
                {
                    if (!cellrec.MatchesCategory(category))
                    {
                        continue;
                    }

                    if (!cellrec.Result.HasValue)
                    {
                        continue;
                    }

                    var sidsrc = new VA.ShapeSheet.SIDSRC((short)shape_id, cellrec.SRC);
                    update.SetFormula(sidsrc, cellrec.Result.Value);
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