using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;

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

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.Color), SrcConstants.CharColor, this.Color);
                yield return CellMetadataItem.Create(nameof(this.Font), SrcConstants.CharFont, this.Font);
                yield return CellMetadataItem.Create(nameof(this.Size), SrcConstants.CharSize, this.Size);
                yield return CellMetadataItem.Create(nameof(this.Style), SrcConstants.CharStyle, this.Style);
                yield return CellMetadataItem.Create(nameof(this.ColorTransparency), SrcConstants.CharColorTransparency, this.ColorTransparency);
                yield return CellMetadataItem.Create(nameof(this.AsianFont), SrcConstants.CharAsianFont, this.AsianFont);
                yield return CellMetadataItem.Create(nameof(this.Case), SrcConstants.CharCase, this.Case);
                yield return CellMetadataItem.Create(nameof(this.ComplexScriptFont), SrcConstants.CharComplexScriptFont, this.ComplexScriptFont);
                yield return CellMetadataItem.Create(nameof(this.ComplexScriptSize), SrcConstants.CharComplexScriptSize, this.ComplexScriptSize);
                yield return CellMetadataItem.Create(nameof(this.DoubleUnderline), SrcConstants.CharDoubleUnderline, this.DoubleUnderline);
                yield return CellMetadataItem.Create(nameof(this.DoubleStrikethrough), SrcConstants.CharDoubleStrikethrough, this.DoubleStrikethrough);
                yield return CellMetadataItem.Create(nameof(this.LangID), SrcConstants.CharLangID, this.LangID);
                yield return CellMetadataItem.Create(nameof(this.FontScale), SrcConstants.CharFontScale, this.FontScale);
                yield return CellMetadataItem.Create(nameof(this.Letterspace), SrcConstants.CharLetterspace, this.Letterspace);
                yield return CellMetadataItem.Create(nameof(this.Locale), SrcConstants.CharLocale, this.Locale);
                yield return CellMetadataItem.Create(nameof(this.LocalizeFont), SrcConstants.CharLocalizeFont, this.LocalizeFont);
                yield return CellMetadataItem.Create(nameof(this.Overline), SrcConstants.CharOverline, this.Overline);
                yield return CellMetadataItem.Create(nameof(this.Perpendicular), SrcConstants.CharPerpendicular, this.Perpendicular);
                yield return CellMetadataItem.Create(nameof(this.Pos), SrcConstants.CharPos, this.Pos);
                yield return CellMetadataItem.Create(nameof(this.RTLText), SrcConstants.CharRTLText, this.RTLText);
                yield return CellMetadataItem.Create(nameof(this.Strikethru), SrcConstants.CharStrikethru, this.Strikethru);
                yield return CellMetadataItem.Create(nameof(this.UseVertical), SrcConstants.CharUseVertical, this.UseVertical);
            }
        }

    }


}