using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue Color { get; set; }
        public VisioAutomation.Core.CellValue Font { get; set; }
        public VisioAutomation.Core.CellValue Size { get; set; }
        public VisioAutomation.Core.CellValue Style { get; set; }
        public VisioAutomation.Core.CellValue ColorTransparency { get; set; }
        public VisioAutomation.Core.CellValue AsianFont { get; set; }
        public VisioAutomation.Core.CellValue Case { get; set; }
        public VisioAutomation.Core.CellValue ComplexScriptFont { get; set; }
        public VisioAutomation.Core.CellValue ComplexScriptSize { get; set; }
        public VisioAutomation.Core.CellValue DoubleStrikethrough { get; set; }
        public VisioAutomation.Core.CellValue DoubleUnderline { get; set; }
        public VisioAutomation.Core.CellValue LangID { get; set; }
        public VisioAutomation.Core.CellValue Locale { get; set; }
        public VisioAutomation.Core.CellValue LocalizeFont { get; set; }
        public VisioAutomation.Core.CellValue Overline { get; set; }
        public VisioAutomation.Core.CellValue Perpendicular { get; set; }
        public VisioAutomation.Core.CellValue Pos { get; set; }
        public VisioAutomation.Core.CellValue RTLText { get; set; }
        public VisioAutomation.Core.CellValue FontScale { get; set; }
        public VisioAutomation.Core.CellValue Letterspace { get; set; }
        public VisioAutomation.Core.CellValue Strikethru { get; set; }
        public VisioAutomation.Core.CellValue UseVertical { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.Color), VisioAutomation.Core.SrcConstants.CharColor, this.Color);
            yield return this.Create(nameof(this.Font), VisioAutomation.Core.SrcConstants.CharFont, this.Font);
            yield return this.Create(nameof(this.Size), VisioAutomation.Core.SrcConstants.CharSize, this.Size);
            yield return this.Create(nameof(this.Style), VisioAutomation.Core.SrcConstants.CharStyle, this.Style);
            yield return this.Create(nameof(this.ColorTransparency), VisioAutomation.Core.SrcConstants.CharColorTransparency,
                this.ColorTransparency);
            yield return this.Create(nameof(this.AsianFont), VisioAutomation.Core.SrcConstants.CharAsianFont, this.AsianFont);
            yield return this.Create(nameof(this.Case), VisioAutomation.Core.SrcConstants.CharCase, this.Case);
            yield return this.Create(nameof(this.ComplexScriptFont), VisioAutomation.Core.SrcConstants.CharComplexScriptFont,
                this.ComplexScriptFont);
            yield return this.Create(nameof(this.ComplexScriptSize), VisioAutomation.Core.SrcConstants.CharComplexScriptSize,
                this.ComplexScriptSize);
            yield return this.Create(nameof(this.DoubleUnderline), VisioAutomation.Core.SrcConstants.CharDoubleUnderline, this.DoubleUnderline);
            yield return this.Create(nameof(this.DoubleStrikethrough), VisioAutomation.Core.SrcConstants.CharDoubleStrikethrough,
                this.DoubleStrikethrough);
            yield return this.Create(nameof(this.LangID), VisioAutomation.Core.SrcConstants.CharLangID, this.LangID);
            yield return this.Create(nameof(this.FontScale), VisioAutomation.Core.SrcConstants.CharFontScale, this.FontScale);
            yield return this.Create(nameof(this.Letterspace), VisioAutomation.Core.SrcConstants.CharLetterspace, this.Letterspace);
            yield return this.Create(nameof(this.Locale), VisioAutomation.Core.SrcConstants.CharLocale, this.Locale);
            yield return this.Create(nameof(this.LocalizeFont), VisioAutomation.Core.SrcConstants.CharLocalizeFont, this.LocalizeFont);
            yield return this.Create(nameof(this.Overline), VisioAutomation.Core.SrcConstants.CharOverline, this.Overline);
            yield return this.Create(nameof(this.Perpendicular), VisioAutomation.Core.SrcConstants.CharPerpendicular, this.Perpendicular);
            yield return this.Create(nameof(this.Pos), VisioAutomation.Core.SrcConstants.CharPos, this.Pos);
            yield return this.Create(nameof(this.RTLText), VisioAutomation.Core.SrcConstants.CharRTLText, this.RTLText);
            yield return this.Create(nameof(this.Strikethru), VisioAutomation.Core.SrcConstants.CharStrikethru, this.Strikethru);
            yield return this.Create(nameof(this.UseVertical), VisioAutomation.Core.SrcConstants.CharUseVertical, this.UseVertical);
        }


        public static List<List<CharacterFormatCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs pairs, VisioAutomation.Core.CellValueType type)
        {
            var reader = CharacterFormatCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, pairs, type);
        }

        public static List<CharacterFormatCells> GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
        {
            var reader = CharacterFormatCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<CharacterFormatCellsBuilder> CharacterFormatCells_lazy_builder = new System.Lazy<CharacterFormatCellsBuilder>();


        class CharacterFormatCellsBuilder : VASS.CellGroups.CellGroupBuilder<Text.CharacterFormatCells>
        {
            public CharacterFormatCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.MultiRow)
            {
            }

            public override Text.CharacterFormatCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new Text.CharacterFormatCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.Color = getcellvalue(nameof(CharacterFormatCells.Color));
                cells.ColorTransparency = getcellvalue(nameof(CharacterFormatCells.ColorTransparency));
                cells.Font = getcellvalue(nameof(CharacterFormatCells.Font));
                cells.Size = getcellvalue(nameof(CharacterFormatCells.Size));
                cells.Style = getcellvalue(nameof(CharacterFormatCells.Style));
                cells.AsianFont = getcellvalue(nameof(CharacterFormatCells.AsianFont));
                cells.AsianFont = getcellvalue(nameof(CharacterFormatCells.AsianFont));
                cells.Case = getcellvalue(nameof(CharacterFormatCells.Case));
                cells.ComplexScriptFont = getcellvalue(nameof(CharacterFormatCells.ComplexScriptFont));
                cells.ComplexScriptSize = getcellvalue(nameof(CharacterFormatCells.ComplexScriptSize));
                cells.DoubleStrikethrough = getcellvalue(nameof(CharacterFormatCells.DoubleStrikethrough));
                cells.DoubleUnderline = getcellvalue(nameof(CharacterFormatCells.DoubleUnderline));
                cells.FontScale = getcellvalue(nameof(CharacterFormatCells.FontScale));
                cells.LangID = getcellvalue(nameof(CharacterFormatCells.LangID));
                cells.Letterspace = getcellvalue(nameof(CharacterFormatCells.Letterspace));
                cells.Locale = getcellvalue(nameof(CharacterFormatCells.Locale));
                cells.LocalizeFont = getcellvalue(nameof(CharacterFormatCells.LocalizeFont));
                cells.Overline = getcellvalue(nameof(CharacterFormatCells.Overline));
                cells.Perpendicular = getcellvalue(nameof(CharacterFormatCells.Perpendicular));
                cells.Pos = getcellvalue(nameof(CharacterFormatCells.Pos));
                cells.RTLText = getcellvalue(nameof(CharacterFormatCells.RTLText));
                cells.Strikethru = getcellvalue(nameof(CharacterFormatCells.Strikethru));
                cells.UseVertical = getcellvalue(nameof(CharacterFormatCells.UseVertical));

                return cells;
            }
        }

    }


}