using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValueLiteral Color { get; set; }
        public VASS.CellValueLiteral Font { get; set; }
        public VASS.CellValueLiteral Size { get; set; }
        public VASS.CellValueLiteral Style { get; set; }
        public VASS.CellValueLiteral ColorTransparency { get; set; }
        public VASS.CellValueLiteral AsianFont { get; set; }
        public VASS.CellValueLiteral Case { get; set; }
        public VASS.CellValueLiteral ComplexScriptFont { get; set; }
        public VASS.CellValueLiteral ComplexScriptSize { get; set; }
        public VASS.CellValueLiteral DoubleStrikethrough { get; set; }
        public VASS.CellValueLiteral DoubleUnderline { get; set; }
        public VASS.CellValueLiteral LangID { get; set; }
        public VASS.CellValueLiteral Locale { get; set; }
        public VASS.CellValueLiteral LocalizeFont { get; set; }
        public VASS.CellValueLiteral Overline { get; set; }
        public VASS.CellValueLiteral Perpendicular { get; set; }
        public VASS.CellValueLiteral Pos { get; set; }
        public VASS.CellValueLiteral RTLText { get; set; }
        public VASS.CellValueLiteral FontScale { get; set; }
        public VASS.CellValueLiteral Letterspace { get; set; }
        public VASS.CellValueLiteral Strikethru { get; set; }
        public VASS.CellValueLiteral UseVertical { get; set; }

        public override IEnumerable<VASS.CellGroups.CellMetadataItem> CellMetadata
        {
            get
            {


                yield return this.Create(nameof(this.Color), VASS.SrcConstants.CharColor, this.Color);
                yield return this.Create(nameof(this.Font), VASS.SrcConstants.CharFont, this.Font);
                yield return this.Create(nameof(this.Size), VASS.SrcConstants.CharSize, this.Size);
                yield return this.Create(nameof(this.Style), VASS.SrcConstants.CharStyle, this.Style);
                yield return this.Create(nameof(this.ColorTransparency), VASS.SrcConstants.CharColorTransparency, this.ColorTransparency);
                yield return this.Create(nameof(this.AsianFont), VASS.SrcConstants.CharAsianFont, this.AsianFont);
                yield return this.Create(nameof(this.Case), VASS.SrcConstants.CharCase, this.Case);
                yield return this.Create(nameof(this.ComplexScriptFont), VASS.SrcConstants.CharComplexScriptFont, this.ComplexScriptFont);
                yield return this.Create(nameof(this.ComplexScriptSize), VASS.SrcConstants.CharComplexScriptSize, this.ComplexScriptSize);
                yield return this.Create(nameof(this.DoubleUnderline), VASS.SrcConstants.CharDoubleUnderline, this.DoubleUnderline);
                yield return this.Create(nameof(this.DoubleStrikethrough), VASS.SrcConstants.CharDoubleStrikethrough, this.DoubleStrikethrough);
                yield return this.Create(nameof(this.LangID), VASS.SrcConstants.CharLangID, this.LangID);
                yield return this.Create(nameof(this.FontScale), VASS.SrcConstants.CharFontScale, this.FontScale);
                yield return this.Create(nameof(this.Letterspace), VASS.SrcConstants.CharLetterspace, this.Letterspace);
                yield return this.Create(nameof(this.Locale), VASS.SrcConstants.CharLocale, this.Locale);
                yield return this.Create(nameof(this.LocalizeFont), VASS.SrcConstants.CharLocalizeFont, this.LocalizeFont);
                yield return this.Create(nameof(this.Overline), VASS.SrcConstants.CharOverline, this.Overline);
                yield return this.Create(nameof(this.Perpendicular), VASS.SrcConstants.CharPerpendicular, this.Perpendicular);
                yield return this.Create(nameof(this.Pos), VASS.SrcConstants.CharPos, this.Pos);
                yield return this.Create(nameof(this.RTLText), VASS.SrcConstants.CharRTLText, this.RTLText);
                yield return this.Create(nameof(this.Strikethru), VASS.SrcConstants.CharStrikethru, this.Strikethru);
                yield return this.Create(nameof(this.UseVertical), VASS.SrcConstants.CharUseVertical, this.UseVertical);
            }
        }


        public static List<List<CharacterFormatCells>> GetCells(IVisio.Page page, ShapeIdPairs pairs, VASS.CellValueType type)
        {
            var reader = CharacterFormatCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, pairs, type);
        }

        public static List<CharacterFormatCells> GetCells(IVisio.Shape shape, VASS.CellValueType type)
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