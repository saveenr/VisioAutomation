using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells : CellGroup
    {
        public CellValueLiteral Color { get; set; }
        public CellValueLiteral Font { get; set; }
        public CellValueLiteral Size { get; set; }
        public CellValueLiteral Style { get; set; }
        public CellValueLiteral ColorTransparency { get; set; }
        public CellValueLiteral AsianFont { get; set; }
        public CellValueLiteral Case { get; set; }
        public CellValueLiteral ComplexScriptFont { get; set; }
        public CellValueLiteral ComplexScriptSize { get; set; }
        public CellValueLiteral DoubleStrikethrough { get; set; }
        public CellValueLiteral DoubleUnderline { get; set; }
        public CellValueLiteral LangID { get; set; }
        public CellValueLiteral Locale { get; set; }
        public CellValueLiteral LocalizeFont { get; set; }
        public CellValueLiteral Overline { get; set; }
        public CellValueLiteral Perpendicular { get; set; }
        public CellValueLiteral Pos { get; set; }
        public CellValueLiteral RTLText { get; set; }
        public CellValueLiteral FontScale { get; set; }
        public CellValueLiteral Letterspace { get; set; }
        public CellValueLiteral Strikethru { get; set; }
        public CellValueLiteral UseVertical { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.CharColor, this.Color);
                yield return SrcValuePair.Create(SrcConstants.CharFont, this.Font);
                yield return SrcValuePair.Create(SrcConstants.CharSize, this.Size);
                yield return SrcValuePair.Create(SrcConstants.CharStyle, this.Style);
                yield return SrcValuePair.Create(SrcConstants.CharColorTransparency, this.ColorTransparency);
                yield return SrcValuePair.Create(SrcConstants.CharAsianFont, this.AsianFont);
                yield return SrcValuePair.Create(SrcConstants.CharCase, this.Case);
                yield return SrcValuePair.Create(SrcConstants.CharComplexScriptFont, this.ComplexScriptFont);
                yield return SrcValuePair.Create(SrcConstants.CharComplexScriptSize, this.ComplexScriptSize);
                yield return SrcValuePair.Create(SrcConstants.CharDoubleUnderline, this.DoubleUnderline);
                yield return SrcValuePair.Create(SrcConstants.CharDoubleStrikethrough, this.DoubleStrikethrough);
                yield return SrcValuePair.Create(SrcConstants.CharLangID, this.LangID);
                yield return SrcValuePair.Create(SrcConstants.CharFontScale, this.FontScale);
                yield return SrcValuePair.Create(SrcConstants.CharLangID, this.LangID);
                yield return SrcValuePair.Create(SrcConstants.CharLetterspace, this.Letterspace);
                yield return SrcValuePair.Create(SrcConstants.CharLocale, this.Locale);
                yield return SrcValuePair.Create(SrcConstants.CharLocalizeFont, this.LocalizeFont);
                yield return SrcValuePair.Create(SrcConstants.CharOverline, this.Overline);
                yield return SrcValuePair.Create(SrcConstants.CharPerpendicular, this.Perpendicular);
                yield return SrcValuePair.Create(SrcConstants.CharPos, this.Pos);
                yield return SrcValuePair.Create(SrcConstants.CharRTLText, this.RTLText);
                yield return SrcValuePair.Create(SrcConstants.CharStrikethru, this.Strikethru);
                yield return SrcValuePair.Create(SrcConstants.CharUseVertical, this.UseVertical);
            }
        }

    }
}