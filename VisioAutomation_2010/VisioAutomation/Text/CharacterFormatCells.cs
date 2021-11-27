using System.Collections.Generic;
using VACG=VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells : VACG.CellGroup
    {
        public Core.CellValue Color { get; set; }
        public Core.CellValue Font { get; set; }
        public Core.CellValue Size { get; set; }
        public Core.CellValue Style { get; set; }
        public Core.CellValue ColorTransparency { get; set; }
        public Core.CellValue AsianFont { get; set; }
        public Core.CellValue Case { get; set; }
        public Core.CellValue ComplexScriptFont { get; set; }
        public Core.CellValue ComplexScriptSize { get; set; }
        public Core.CellValue DoubleStrikethrough { get; set; }
        public Core.CellValue DoubleUnderline { get; set; }
        public Core.CellValue LangID { get; set; }
        public Core.CellValue Locale { get; set; }
        public Core.CellValue LocalizeFont { get; set; }
        public Core.CellValue Overline { get; set; }
        public Core.CellValue Perpendicular { get; set; }
        public Core.CellValue Pos { get; set; }
        public Core.CellValue RTLText { get; set; }
        public Core.CellValue FontScale { get; set; }
        public Core.CellValue Letterspace { get; set; }
        public Core.CellValue Strikethru { get; set; }
        public Core.CellValue UseVertical { get; set; }

        public override IEnumerable<VACG.CellMetadataItem> GetCellMetadata()
        {
            yield return this._create(nameof(this.Color), Core.SrcConstants.CharColor, this.Color);
            yield return this._create(nameof(this.Font), Core.SrcConstants.CharFont, this.Font);
            yield return this._create(nameof(this.Size), Core.SrcConstants.CharSize, this.Size);
            yield return this._create(nameof(this.Style), Core.SrcConstants.CharStyle, this.Style);
            yield return this._create(nameof(this.ColorTransparency), Core.SrcConstants.CharColorTransparency,
                this.ColorTransparency);
            yield return this._create(nameof(this.AsianFont), Core.SrcConstants.CharAsianFont, this.AsianFont);
            yield return this._create(nameof(this.Case), Core.SrcConstants.CharCase, this.Case);
            yield return this._create(nameof(this.ComplexScriptFont), Core.SrcConstants.CharComplexScriptFont,
                this.ComplexScriptFont);
            yield return this._create(nameof(this.ComplexScriptSize), Core.SrcConstants.CharComplexScriptSize,
                this.ComplexScriptSize);
            yield return this._create(nameof(this.DoubleUnderline), Core.SrcConstants.CharDoubleUnderline, this.DoubleUnderline);
            yield return this._create(nameof(this.DoubleStrikethrough), Core.SrcConstants.CharDoubleStrikethrough,
                this.DoubleStrikethrough);
            yield return this._create(nameof(this.LangID), Core.SrcConstants.CharLangID, this.LangID);
            yield return this._create(nameof(this.FontScale), Core.SrcConstants.CharFontScale, this.FontScale);
            yield return this._create(nameof(this.Letterspace), Core.SrcConstants.CharLetterspace, this.Letterspace);
            yield return this._create(nameof(this.Locale), Core.SrcConstants.CharLocale, this.Locale);
            yield return this._create(nameof(this.LocalizeFont), Core.SrcConstants.CharLocalizeFont, this.LocalizeFont);
            yield return this._create(nameof(this.Overline), Core.SrcConstants.CharOverline, this.Overline);
            yield return this._create(nameof(this.Perpendicular), Core.SrcConstants.CharPerpendicular, this.Perpendicular);
            yield return this._create(nameof(this.Pos), Core.SrcConstants.CharPos, this.Pos);
            yield return this._create(nameof(this.RTLText), Core.SrcConstants.CharRTLText, this.RTLText);
            yield return this._create(nameof(this.Strikethru), Core.SrcConstants.CharStrikethru, this.Strikethru);
            yield return this._create(nameof(this.UseVertical), Core.SrcConstants.CharUseVertical, this.UseVertical);
        }


        public static List<List<CharacterFormatCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs pairs, Core.CellValueType type)
        {
            var reader = CharacterFormatCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, pairs, type);
        }

        public static List<CharacterFormatCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = CharacterFormatCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<Builder> CharacterFormatCells_lazy_builder = new System.Lazy<Builder>();


        class Builder : VACG.CellGroupBuilder<CharacterFormatCells>
        {
            public Builder() : base(VACG.CellGroupBuilderType.MultiRow)
            {
            }

            public override CharacterFormatCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new CharacterFormatCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.Color = getcellvalue(nameof(Color));
                cells.ColorTransparency = getcellvalue(nameof(ColorTransparency));
                cells.Font = getcellvalue(nameof(Font));
                cells.Size = getcellvalue(nameof(Size));
                cells.Style = getcellvalue(nameof(Style));
                cells.AsianFont = getcellvalue(nameof(AsianFont));
                cells.AsianFont = getcellvalue(nameof(AsianFont));
                cells.Case = getcellvalue(nameof(Case));
                cells.ComplexScriptFont = getcellvalue(nameof(ComplexScriptFont));
                cells.ComplexScriptSize = getcellvalue(nameof(ComplexScriptSize));
                cells.DoubleStrikethrough = getcellvalue(nameof(DoubleStrikethrough));
                cells.DoubleUnderline = getcellvalue(nameof(DoubleUnderline));
                cells.FontScale = getcellvalue(nameof(FontScale));
                cells.LangID = getcellvalue(nameof(LangID));
                cells.Letterspace = getcellvalue(nameof(Letterspace));
                cells.Locale = getcellvalue(nameof(Locale));
                cells.LocalizeFont = getcellvalue(nameof(LocalizeFont));
                cells.Overline = getcellvalue(nameof(Overline));
                cells.Perpendicular = getcellvalue(nameof(Perpendicular));
                cells.Pos = getcellvalue(nameof(Pos));
                cells.RTLText = getcellvalue(nameof(RTLText));
                cells.Strikethru = getcellvalue(nameof(Strikethru));
                cells.UseVertical = getcellvalue(nameof(UseVertical));

                return cells;
            }
        }

    }


}