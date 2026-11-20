using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page shapesheet/query-the-shapesheet.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ShapesheetQueryTheShapesheetSamples
    {
        public static void Sample_01(IVisio.Shape shape, VisioAutomation.Core.Src src_fg)
        {
            // <sample id="shapesheet/query-the-shapesheet#01" doc="shapesheet/query-the-shapesheet.md" line="43">
            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_fg = query.Columns.Add(src_fg);
            var formulas = query.GetFormulas(shape);     // Data.DataRows<string>
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape, VisioAutomation.Core.Src src_fg)
        {
            // setup: the code of block shapesheet/query-the-shapesheet#01, which this block continues
            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_fg = query.Columns.Add(src_fg);
            var formulas = query.GetFormulas(shape);     // Data.DataRows<string>

            // <sample id="shapesheet/query-the-shapesheet#02" doc="shapesheet/query-the-shapesheet.md" line="51">
            string f_fg = formulas[0][0];
            string f_fg_alt = formulas[0][col_fg.Ordinal];
            // </sample>
        }

        public static void Sample_03(IVisio.Shape shape, VisioAutomation.Core.Src src_fg, VisioAutomation.Core.Src src_bg, VisioAutomation.Core.Src src_filpat)
        {
            // <sample id="shapesheet/query-the-shapesheet#03" doc="shapesheet/query-the-shapesheet.md" line="60">
            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_fg = query.Columns.Add(src_fg);
            var col_bg = query.Columns.Add(src_bg);
            var col_filpat = query.Columns.Add(src_filpat);
            var formulas = query.GetFormulas(shape);

            string f_fg = formulas[0][col_fg.Ordinal];
            string f_bg = formulas[0][col_bg.Ordinal];
            string f_filpat = formulas[0][col_filpat.Ordinal];
            // </sample>
        }

        public static void Sample_04(IVisio.Shape shape, VA.ShapeSheet.Query.CellQuery query)
        {
            // <sample id="shapesheet/query-the-shapesheet#04" doc="shapesheet/query-the-shapesheet.md" line="76">
            VA.ShapeSheet.Data.DataRows<double> results_double = query.GetResults<double>(shape);
            VA.ShapeSheet.Data.DataRows<int>    results_int    = query.GetResults<int>(shape);
            VA.ShapeSheet.Data.DataRows<bool>   results_bool   = query.GetResults<bool>(shape);
            VA.ShapeSheet.Data.DataRows<string> results_string = query.GetResults<string>(shape);
            // </sample>
        }

        public static void Sample_05(IVisio.Shape shape1)
        {
            // <sample id="shapesheet/query-the-shapesheet#05" doc="shapesheet/query-the-shapesheet.md" line="87">
            var query = new VA.ShapeSheet.Query.SectionQuery();
            var sec_cols = query.Add(IVisio.VisSectionIndices.visSectionProp);
            sec_cols.Add(VA.Core.SrcConstants.CharColor);
            var formulas = query.GetFormulas(shape1);     // Data.DataRowGroup<string>
            // </sample>
        }

        public static void Sample_06(IVisio.Shape shape1)
        {
            // setup: the code of block shapesheet/query-the-shapesheet#05, which this block continues
            var query = new VA.ShapeSheet.Query.SectionQuery();
            var sec_cols = query.Add(IVisio.VisSectionIndices.visSectionProp);
            sec_cols.Add(VA.Core.SrcConstants.CharColor);
            var formulas = query.GetFormulas(shape1);     // Data.DataRowGroup<string>

            // <sample id="shapesheet/query-the-shapesheet#06" doc="shapesheet/query-the-shapesheet.md" line="96">
            foreach (var section in formulas)
            {
                foreach (var row in section)
                {
                    System.Console.WriteLine(row[0]);    // first column of this row
                }
            }
            // </sample>
        }

        public static void Sample_07(IVisio.Page page, IVisio.Shape shape1, IVisio.Shape shape2, VisioAutomation.Core.Src src_fg, VisioAutomation.Core.Src src_bg)
        {
            // <sample id="shapesheet/query-the-shapesheet#07" doc="shapesheet/query-the-shapesheet.md" line="114">
            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_fg = query.Columns.Add(src_fg);
            var col_bg = query.Columns.Add(src_bg);
            var shapeids = new[] { shape1.ID, shape2.ID };
            var formulas = query.GetFormulas(page, shapeids);     // Data.DataRows<string>
            // </sample>
        }

        public static void Sample_08(IVisio.Shape s1, IVisio.Shape s2, IVisio.Page page1, IVisio.Shape s3, IVisio.Shape s4)
        {
            // <sample id="shapesheet/query-the-shapesheet#08" doc="shapesheet/query-the-shapesheet.md" line="126">
            var query = new VA.ShapeSheet.Query.SectionQuery();
            var sec_cols = query.Add(IVisio.VisSectionIndices.visSectionProp);
            sec_cols.Add(VA.Core.SrcConstants.CharColor);
            var pairs = VA.Core.ShapeIDPairs.FromShapes(s1, s2, s3, s4);
            var formulas = query.GetFormulas(page1, pairs);       // Data.DataRowGroups<string>
            // </sample>
        }
    }
}
