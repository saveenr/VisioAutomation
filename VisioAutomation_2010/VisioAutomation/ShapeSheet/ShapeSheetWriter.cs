using VisioAutomation.ShapeSheet.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public class ShapeSheetWriter
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        private WriterCollection<SidSrc> FormulaRecords_SidSrc;
        private WriterCollection<Src> FormulaRecords_Src;
        private WriterCollection<Src> ResultRecords_Src;
        private WriterCollection<SidSrc> ResultRecords_SidSrc;

        public ShapeSheetWriter()
        {
        }

        public void Clear()
        {
            FormulaRecords_SidSrc?.Clear();
            FormulaRecords_Src?.Clear();
            ResultRecords_Src?.Clear();
            ResultRecords_SidSrc?.Clear();
        }

        protected IVisio.VisGetSetArgs ComputeGetResultFlags()
        {
            var flags = this.combine_blastguards_and_testcircular_flags();

            flags |= IVisio.VisGetSetArgs.visGetStrings;

            return flags;
        }

        protected IVisio.VisGetSetArgs ComputeGetFormulaFlags()
        {
            var common_flags = this.combine_blastguards_and_testcircular_flags();
            var formula_flags = (short)IVisio.VisGetSetArgs.visSetUniversalSyntax;
            var combined_flags = (short)common_flags | formula_flags;
            return (IVisio.VisGetSetArgs)combined_flags;
        }

        private IVisio.VisGetSetArgs combine_blastguards_and_testcircular_flags()
        {
            var f_bg = this.BlastGuards ? IVisio.VisGetSetArgs.visSetBlastGuards : 0;
            var f_tc = this.TestCircular ? IVisio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = ((short)f_bg) | ((short)f_tc);
            return (IVisio.VisGetSetArgs)flags;
        }

        public void Commit(IVisio.Shape shape)
        {
            var surface = new ShapeSheetSurface(shape);
            this.Commit(surface);
        }

        public void Commit(IVisio.Page page)
        {
            var surface = new ShapeSheetSurface(page);
            this.Commit(surface);
        }

        public void Commit(VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            this.CommitFormulaRecordsByType(surface, Streams.StreamType.Src);
            this.CommitFormulaRecordsByType(surface, Streams.StreamType.SidSrc);
            this.CommitResultRecordsByType(surface, Streams.StreamType.Src);
            this.CommitResultRecordsByType(surface, Streams.StreamType.SidSrc);
        }

        public void SetFormula(Src src, CellValueLiteral formula)
        {
            this.__SetFormulaIgnoreNull(src, formula);
        }

        public void SetFormula(short id, Src src, CellValueLiteral formula)
        {
            var sidsrc = new SidSrc(id, src);
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        public void SetFormula(SidSrc sidsrc, CellValueLiteral formula)
        {
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        private void __SetFormulaIgnoreNull(Src src, CellValueLiteral formula)
        {
            if (this.FormulaRecords_Src == null)
            {
                this.FormulaRecords_Src = new WriterCollection<Src>();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_Src.Add(src,formula.Value);
            }
        }

        private void __SetFormulaIgnoreNull(SidSrc sidsrc, CellValueLiteral formula)
        {
            if (this.FormulaRecords_SidSrc == null)
            {
                this.FormulaRecords_SidSrc = new WriterCollection<SidSrc>();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_SidSrc.Add(sidsrc, formula.Value);
            }
        }

        private VisioAutomation.ShapeSheet.Streams.StreamArray buildstream_sidsrc(WriterCollection<SidSrc> wcs)
        {
            var builder = new VisioAutomation.ShapeSheet.Streams.FixedSidSrcStreamBuilder(wcs.Count);
            builder.AddRange(wcs.EnumCoords());
            return builder.ToStream();
        }

        private VisioAutomation.ShapeSheet.Streams.StreamArray buildstream_src(WriterCollection<Src> wcs)
        {
            var builder = new VisioAutomation.ShapeSheet.Streams.FixedSrcStreamBuilder(wcs.Count);
            builder.AddRange(wcs.EnumCoords());
            return builder.ToStream();
        }

        private void CommitFormulaRecordsByType(ShapeSheetSurface surface, Streams.StreamType cell_coord)
        {
            if (cell_coord == Streams.StreamType.SidSrc && (this.FormulaRecords_SidSrc == null || this.FormulaRecords_SidSrc.Count <1))
            {
                return;
            }

            if (cell_coord == Streams.StreamType.Src && (this.FormulaRecords_Src == null || this.FormulaRecords_Src.Count <1))
            {
                return;
            }

            var stream = cell_coord == Streams.StreamType.SidSrc ? this.buildstream_sidsrc(this.FormulaRecords_SidSrc) : this.buildstream_src(this.FormulaRecords_Src);
            var formulas = cell_coord == Streams.StreamType.SidSrc ? this.FormulaRecords_SidSrc.BuildValues() : this.FormulaRecords_Src.BuildValues();

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetFormulaFlags();

            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        public void SetResult(Src src, CellValueLiteral result)
        {
            if (this.ResultRecords_Src == null)
            {
                this.ResultRecords_Src = new WriterCollection<Src>();
            }

            this.ResultRecords_Src.Add(src,result.Value);
        }

        public void SetResult(short id, Src src, CellValueLiteral result)
        {
            var sidsrc = new SidSrc(id, src);
            this.SetResult(sidsrc, result.Value);
        }

        public void SetResult(SidSrc sidsrc, CellValueLiteral result)
        {
            if (this.ResultRecords_SidSrc == null)
            {
                this.ResultRecords_SidSrc = new WriterCollection<SidSrc>();
            }

            this.ResultRecords_SidSrc.Add(sidsrc, result.Value);
        }

        private void CommitResultRecordsByType(ShapeSheetSurface surface, Streams.StreamType cell_coord)
        {
            if (cell_coord == Streams.StreamType.SidSrc && (this.ResultRecords_SidSrc == null || this.ResultRecords_SidSrc.Count < 1))
            {
                return;
            }

            if (cell_coord == Streams.StreamType.Src && (this.ResultRecords_Src == null || this.ResultRecords_Src.Count <1))
            {
                return;
            }

            var stream = cell_coord == Streams.StreamType.SidSrc ? this.buildstream_sidsrc(this.ResultRecords_SidSrc) : this.buildstream_src(this.ResultRecords_Src);
            var results = cell_coord == Streams.StreamType.SidSrc ? this.ResultRecords_SidSrc.BuildValues(): this.ResultRecords_Src.BuildValues();
            const object[] unitcodes = null;
            
            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetResultFlags();
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }

















    public class ShapeSheetWriterSrc
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        private WriterCollection<Src> FormulaRecords_Src;
        private WriterCollection<Src> ResultRecords_Src;

        public ShapeSheetWriterSrc()
        {
        }

        public void Clear()
        {
            FormulaRecords_Src?.Clear();
            ResultRecords_Src?.Clear();
        }

        protected IVisio.VisGetSetArgs ComputeGetResultFlags()
        {
            var flags = this.combine_blastguards_and_testcircular_flags();

            flags |= IVisio.VisGetSetArgs.visGetStrings;

            return flags;
        }

        protected IVisio.VisGetSetArgs ComputeGetFormulaFlags()
        {
            var common_flags = this.combine_blastguards_and_testcircular_flags();
            var formula_flags = (short)IVisio.VisGetSetArgs.visSetUniversalSyntax;
            var combined_flags = (short)common_flags | formula_flags;
            return (IVisio.VisGetSetArgs)combined_flags;
        }

        private IVisio.VisGetSetArgs combine_blastguards_and_testcircular_flags()
        {
            var f_bg = this.BlastGuards ? IVisio.VisGetSetArgs.visSetBlastGuards : 0;
            var f_tc = this.TestCircular ? IVisio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = ((short)f_bg) | ((short)f_tc);
            return (IVisio.VisGetSetArgs)flags;
        }

        public void Commit(IVisio.Shape shape)
        {
            var surface = new ShapeSheetSurface(shape);
            this.Commit(surface);
        }

        public void Commit(IVisio.Page page)
        {
            var surface = new ShapeSheetSurface(page);
            this.Commit(surface);
        }

        public void Commit(VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            this.CommitFormulaRecordsByType(surface);
            this.CommitResultRecordsByType(surface);
        }

        public void SetFormula(Src src, CellValueLiteral formula)
        {
            this.__SetFormulaIgnoreNull(src, formula);
        }
        
        private void __SetFormulaIgnoreNull(Src src, CellValueLiteral formula)
        {
            if (this.FormulaRecords_Src == null)
            {
                this.FormulaRecords_Src = new WriterCollection<Src>();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_Src.Add(src, formula.Value);
            }
        }

        private VisioAutomation.ShapeSheet.Streams.StreamArray buildstream_src(WriterCollection<Src> wcs)
        {
            var builder = new VisioAutomation.ShapeSheet.Streams.FixedSrcStreamBuilder(wcs.Count);
            builder.AddRange(wcs.EnumCoords());
            return builder.ToStream();
        }

        private void CommitFormulaRecordsByType(ShapeSheetSurface surface)
        {
            if ((this.FormulaRecords_Src == null || this.FormulaRecords_Src.Count < 1))
            {
                return;
            }

            var stream = this.buildstream_src(this.FormulaRecords_Src);
            var formulas = this.FormulaRecords_Src.BuildValues();

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetFormulaFlags();

            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        public void SetResult(Src src, CellValueLiteral result)
        {
            if (this.ResultRecords_Src == null)
            {
                this.ResultRecords_Src = new WriterCollection<Src>();
            }

            this.ResultRecords_Src.Add(src, result.Value);
        }

        private void CommitResultRecordsByType(ShapeSheetSurface surface)
        {
            if (this.ResultRecords_Src == null || this.ResultRecords_Src.Count < 1)
            {
                return;
            }

            var stream = this.buildstream_src(this.ResultRecords_Src);
            var results = this.ResultRecords_Src.BuildValues();
            const object[] unitcodes = null;

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetResultFlags();
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }







    public class ShapeSheetWriterSidSrc
    {
        public bool BlastGuards { get; set; }
        public bool TestCircular { get; set; }

        private WriterCollection<SidSrc> FormulaRecords_SidSrc;
        private WriterCollection<SidSrc> ResultRecords_SidSrc;

        public ShapeSheetWriterSidSrc()
        {
        }

        public void Clear()
        {
            FormulaRecords_SidSrc?.Clear();
            ResultRecords_SidSrc?.Clear();
        }

        protected IVisio.VisGetSetArgs ComputeGetResultFlags()
        {
            var flags = this.combine_blastguards_and_testcircular_flags();

            flags |= IVisio.VisGetSetArgs.visGetStrings;

            return flags;
        }

        protected IVisio.VisGetSetArgs ComputeGetFormulaFlags()
        {
            var common_flags = this.combine_blastguards_and_testcircular_flags();
            var formula_flags = (short)IVisio.VisGetSetArgs.visSetUniversalSyntax;
            var combined_flags = (short)common_flags | formula_flags;
            return (IVisio.VisGetSetArgs)combined_flags;
        }

        private IVisio.VisGetSetArgs combine_blastguards_and_testcircular_flags()
        {
            var f_bg = this.BlastGuards ? IVisio.VisGetSetArgs.visSetBlastGuards : 0;
            var f_tc = this.TestCircular ? IVisio.VisGetSetArgs.visSetTestCircular : 0;

            var flags = ((short)f_bg) | ((short)f_tc);
            return (IVisio.VisGetSetArgs)flags;
        }

        public void Commit(IVisio.Shape shape)
        {
            var surface = new ShapeSheetSurface(shape);
            this.Commit(surface);
        }

        public void Commit(IVisio.Page page)
        {
            var surface = new ShapeSheetSurface(page);
            this.Commit(surface);
        }

        public void Commit(VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            this.CommitFormulaRecordsByType(surface);
            this.CommitResultRecordsByType(surface);
        }

        public void SetFormula(short id, Src src, CellValueLiteral formula)
        {
            var sidsrc = new SidSrc(id, src);
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        public void SetFormula(SidSrc sidsrc, CellValueLiteral formula)
        {
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        private void __SetFormulaIgnoreNull(SidSrc sidsrc, CellValueLiteral formula)
        {
            if (this.FormulaRecords_SidSrc == null)
            {
                this.FormulaRecords_SidSrc = new WriterCollection<SidSrc>();
            }

            if (formula.HasValue)
            {
                this.FormulaRecords_SidSrc.Add(sidsrc, formula.Value);
            }
        }

        private VisioAutomation.ShapeSheet.Streams.StreamArray buildstream_sidsrc(WriterCollection<SidSrc> wcs)
        {
            var builder = new VisioAutomation.ShapeSheet.Streams.FixedSidSrcStreamBuilder(wcs.Count);
            builder.AddRange(wcs.EnumCoords());
            return builder.ToStream();
        }

        private void CommitFormulaRecordsByType(ShapeSheetSurface surface)
        {
            if ((this.FormulaRecords_SidSrc == null || this.FormulaRecords_SidSrc.Count < 1))
            {
                return;
            }

            var stream = this.buildstream_sidsrc(this.FormulaRecords_SidSrc);
            var formulas = this.FormulaRecords_SidSrc.BuildValues();

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetFormulaFlags();

            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        public void SetResult(short id, Src src, CellValueLiteral result)
        {
            var sidsrc = new SidSrc(id, src);
            this.SetResult(sidsrc, result.Value);
        }

        public void SetResult(SidSrc sidsrc, CellValueLiteral result)
        {
            if (this.ResultRecords_SidSrc == null)
            {
                this.ResultRecords_SidSrc = new WriterCollection<SidSrc>();
            }

            this.ResultRecords_SidSrc.Add(sidsrc, result.Value);
        }

        private void CommitResultRecordsByType(ShapeSheetSurface surface)
        {
            if ((this.ResultRecords_SidSrc == null || this.ResultRecords_SidSrc.Count < 1))
            {
                return;
            }

            var stream = this.buildstream_sidsrc(this.ResultRecords_SidSrc);
            var results = this.ResultRecords_SidSrc.BuildValues();
            const object[] unitcodes = null;

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetResultFlags();
            surface.SetResults(stream, unitcodes, results, (short)flags);
        }
    }

}