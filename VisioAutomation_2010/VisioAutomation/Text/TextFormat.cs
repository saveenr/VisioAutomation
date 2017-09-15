using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextFormat
    {
        public List<CharacterFormatCells> CharacterFormats { get; private set; }
        public List<ParagraphFormatCells> ParagraphFormats { get; private set; }
        public TextBlockCells TextBlock { get; private set; }
        public TextXFormCells TextXForm { get; private set; }
        public List<TextRun> CharacterTextRuns { get; private set; }
        public List<TextRun> ParagraphTextRuns { get; private set; }
        public List<TabStop> TabStops { get; private set; }

        private static List<TextRun> GetTextRuns(
            IVisio.Shape shape,
            IVisio.VisRunTypes runtype,
            bool collect_text)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var runs = new List<TextRun>();

            // based on this example: http://blogs.msdn.com/visio/archive/2006/08/18/704811.aspx
            // Get the Characters object representing the shape text
            var chars = shape.Characters;
            int num_chars = chars.CharCount;
            int run_end = 1;

            int index = 0;

            // Find the beginning point and end point of every text run in the shape
            for (int c = 0; c < num_chars; c = run_end)
            {
                // Set the begin and end of the Characters object to the current position
                chars.Begin = c;
                chars.End = c + 1;

                // Get the beginning and end of this character run
                int run_begin = chars.RunBegin[(short)runtype];
                run_end = chars.RunEnd[(short)runtype];

                // Set the begin and end of the Characters object to this run
                chars.Begin = run_begin;
                chars.End = run_end;

                // Record the text in this run
                string t = null;
                if (collect_text)
                {
                    t = chars.TextAsString;
                }

                var textrun = new TextRun(index, run_begin, run_end, t);
                index++;
                runs.Add(textrun);

                // As the for loop proceeds, c is set to the end of the current run
            }

            return runs;
        }
        
        public static TextFormat GetFormat(IVisio.Shape shape, CellValueType cvt)
        {
            var cells = new TextFormat();
            cells.CharacterFormats = CharacterFormatCells.GetCells(shape, cvt);
            cells.ParagraphFormats = ParagraphFormatCells.GetCells(shape, cvt);
            cells.TextBlock = TextBlockCells.GetCells(shape, cvt);
            if (HasTextXFormCells(shape))
            {
                cells.TextXForm = TextXFormCells.GetCells(shape, cvt);
            }
            cells.CharacterTextRuns = TextFormat.GetTextRuns(shape, IVisio.VisRunTypes.visCharPropRow, true);
            cells.ParagraphTextRuns = TextFormat.GetTextRuns(shape, IVisio.VisRunTypes.visParaPropRow, true);
            cells.TabStops = TextHelper.GetTabStops(shape);
            return cells;
        }

        public static bool HasTextXFormCells(IVisio.Shape shape)
        {
            return (
                shape.RowExists[
                    (short) IVisio.VisSectionIndices.visSectionObject, 
                    (short) IVisio.VisRowIndices.visRowTextXForm,
                    (short) 0] != 0) ;
        }

        public static List<TextFormat> GetFormat(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var charcells = CharacterFormatCells.GetCells(page, shapeids, cvt);
            var paracells = ParagraphFormatCells.GetCells(page, shapeids, cvt);
            var textblockcells = TextBlockCells.GetCells(page, shapeids, cvt);
            var page_shapes = page.Shapes;
            var formats = new List<TextFormat>(shapeids.Count);
            for (int i = 0; i < shapeids.Count; i++)
            {
                var format = new TextFormat();
                format.CharacterFormats = charcells[i];
                format.ParagraphFormats = paracells[i];
                format.TextBlock = textblockcells[i];
                formats.Add(format);

                var shape = page_shapes.ItemFromID[shapeids[i]];
                format.CharacterTextRuns = TextFormat.GetTextRuns(shape, IVisio.VisRunTypes.visCharPropRow, true);
                format.ParagraphTextRuns = TextFormat.GetTextRuns(shape, IVisio.VisRunTypes.visParaPropRow, true);

                format.TabStops = TextHelper.GetTabStops(shape);
            }

            return formats;
        }
    }
}