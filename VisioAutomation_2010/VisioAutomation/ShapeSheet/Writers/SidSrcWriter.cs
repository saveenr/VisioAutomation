﻿using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class SidSrcWriter : WriterBase
    {
        public SidSrcWriter() : base(CellCoordinateType.SidSrc)
        {
        }

        public void CommitFormulas(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            this.CommitFormulas(surface);
        }

        public void CommitFormulas(IVisio.Page page)
        {
            var surface = new SurfaceTarget(page);
            this.CommitFormulas(surface);
        }

        public void CommitResults(IVisio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            this.CommitResults(surface);
        }

        public void CommitResults(IVisio.Page page)
        {
            var surface = new SurfaceTarget(page);
            this.CommitResults(surface);
        }

        
        public void SetValue(short id, Src src, CellValueLiteral formula)
        {
            var sidsrc = new SidSrc(id, src);
            this.__SetValueIgnoreNull(sidsrc, formula);
        }

        public void SetValue(SidSrc sidsrc, CellValueLiteral formula)
        {
            this.__SetValueIgnoreNull(sidsrc, formula);
        }

        public void SetValues(short id, CellGroups.CellGroup cgb, short row)
        {
            var pairs = cgb.SidSrcValuePairs_NewRow(id, row);
            foreach (var pair in pairs)
            {
                this.SetValue(pair.ShapeID, pair.Src, pair.Value);
            }
        }

        public void SetValues(short id, CellGroups.CellGroup cgb)
        {
            foreach (var pair in cgb.SrcValuePairs)
            {
                this.SetValue(id, pair.Src, pair.Value);
            }
        }

        private void __SetValueIgnoreNull(SidSrc sidsrc, CellValueLiteral formula)
        {
            if (this._records == null)
            {
                this._records = new WriteRecordList(CellCoordinateType.SidSrc);
            }

            if (formula.HasValue)
            {
                this._records.Add(sidsrc, formula.Value);
            }
        }

        public void CommitFormulas(SurfaceTarget surface)
        {
            if ((this._records == null || this._records.Count < 1))
            {
                return;
            }

            var stream = this._records.BuildSidSrcStream();
            var formulas = this._records.BuildValuesArray();

            if (stream.Array.Length == 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var flags = this.ComputeGetFormulaFlags();

            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        public void CommitResults(SurfaceTarget surface)
        {
            if ((this._records == null || this._records.Count < 1))
            {
                return;
            }

            var stream = this._records.BuildSidSrcStream();
            var results = this._records.BuildValuesArray();
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