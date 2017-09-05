using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionsQuery
    {
        public SectionQueryList SectionQueries { get; }

        private SectionInfoCache _cache;

        public SectionsQuery()
        {
            this.SectionQueries = new SectionQueryList(0);
        }

        private static void RestrictToShapesOnly(SurfaceTarget surface)
        {
            if (surface.Shape == null)
            {
                string msg = "Target must be Shape not Page or Master";
                throw new System.ArgumentException(msg);
            }
        }

        public SectionsQueryOutput<string> GetFormulas(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetFormulas(surface);
        }

        public SectionsQueryOutput<string> GetFormulas(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var shapes = new List<Microsoft.Office.Interop.Visio.Shape> { surface.Shape };

            this.cache_section_info(shapes);
            var srcstream = this._build_src_stream();
            var values = surface.GetFormulasU(srcstream);
            var shape_index = 0;
            var sectioninfo = this.GetSectionInfoForShape(shape_index, _cache);
            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<string>(values);
            var output_for_shape = this._create_output_for_shape(surface.ID16, sectioninfo, seg_builder);

            return output_for_shape;
        }

        public SectionsQueryOutput<TResult> GetResults<TResult>(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetResults<TResult>(surface);
        }

        public SectionsQueryOutput<TResult> GetResults<TResult>(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var shapes = new List<Microsoft.Office.Interop.Visio.Shape> { surface.Shape };

            this.cache_section_info(shapes);
            var srcstream = this._build_src_stream();
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var shape_index = 0;
            var sectioninfo = this.GetSectionInfoForShape(shape_index, _cache);
            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<TResult>(values);
            var output_for_shape = this._create_output_for_shape(surface.ID16, sectioninfo, seg_builder);
            return output_for_shape;
        }

        public SectionsQueryOutput<ShapeSheet.CellData> GetFormulasAndResults(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return this.GetFormulasAndResults(surface);
        }

        public SectionsQueryOutput<ShapeSheet.CellData> GetFormulasAndResults(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var shapes = new List<Microsoft.Office.Interop.Visio.Shape> { surface.Shape };

            this.cache_section_info(shapes);
            var srcstream = this._build_src_stream();
            const object[] unitcodes = null;
            var formulas = surface.GetFormulasU(srcstream);
            var results = surface.GetResults<string>(srcstream, unitcodes);
            var combined_data = QueryUtil._combine_formulas_and_results(formulas, results);

            var shape_index = 0;
            var sectioninfo = this.GetSectionInfoForShape(shape_index, _cache);
            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<CellData>(combined_data);
            var output_for_shape = this._create_output_for_shape(surface.ID16, sectioninfo, seg_builder);
            return output_for_shape;
        }

        public SectionsQueryOutputList<string> GetFormulas(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetFormulas(surface, shapeids);
        }

        public SectionsQueryOutputList<string> GetFormulas(SurfaceTarget surface, IList<int> shapeids)
        {
            var shapes = new List<Microsoft.Office.Interop.Visio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Shapes.ItemFromID16[(short)shapeid]));

            this.cache_section_info(shapes);
            var srcstream = this._build_sidsrc_stream(shapeids);
            var values = surface.GetFormulasU(srcstream);
            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<string>(values);
            var list = this._create_outputs_for_shapes(shapeids, _cache, seg_builder);
            return list;
        }

        public SectionsQueryOutputList<TResult> GetResults<TResult>(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetResults<TResult>(surface, shapeids);
        }

        public SectionsQueryOutputList<TResult> GetResults<TResult>(SurfaceTarget surface, IList<int> shapeids)
        {
            var shapes = new List<Microsoft.Office.Interop.Visio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Shapes.ItemFromID16[(short)shapeid]));

            this.cache_section_info(shapes);
            var srcstream = this._build_sidsrc_stream(shapeids);
            const object[] unitcodes = null;
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<TResult>(values);
            var list = this._create_outputs_for_shapes(shapeids, _cache, seg_builder);
            return list;
        }

        public SectionsQueryOutputList<ShapeSheet.CellData> GetFormulasAndResults(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetFormulasAndResults(surface, shapeids);
        }

        public SectionsQueryOutputList<ShapeSheet.CellData> GetFormulasAndResults(SurfaceTarget surface, IList<int> shapeids)
        {
            var shapes = new List<Microsoft.Office.Interop.Visio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Shapes.ItemFromID16[(short)shapeid]));

            this.cache_section_info(shapes);
            var srcstream = this._build_sidsrc_stream(shapeids);
            const object[] unitcodes = null;
            var results = surface.GetResults<string>(srcstream, unitcodes);
            var formulas = surface.GetFormulasU(srcstream);
            var combined_data = QueryUtil._combine_formulas_and_results(formulas, results);

            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<CellData>(combined_data);
            var r = this._create_outputs_for_shapes(shapeids, _cache, seg_builder);
            return r;
        }

        private SectionsQueryOutputList<T> _create_outputs_for_shapes<T>(IList<int> shapeids, SectionInfoCache cache, VisioAutomation.Utilities.ArraySegmentReader<T> segReader)
        {
            var output_for_all_shapes = new SectionsQueryOutputList<T>();

            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shapeid = shapeids[shape_index];
                var secinfo = this.GetSectionInfoForShape(shape_index, cache);
                var output_for_shape = this._create_output_for_shape((short)shapeid, secinfo, segReader);
                output_for_all_shapes.Add(output_for_shape);
            }

            return output_for_all_shapes;
        }

        private List<SectionInfo> GetSectionInfoForShape(int shape_index, SectionInfoCache cache)
        {
            if (cache.CountShapes > 0)
            {
                return cache.GetSectionInfosForShapeAtIndex(shape_index);
            }
            return null;
        }

        private SectionsQueryOutput<T> _create_output_for_shape<T>(short shapeid, List<SectionInfo> section_infos, VisioAutomation.Utilities.ArraySegmentReader<T> segReader)
        {
            int original_seg_size = segReader.Count;

            int results_cell_count = 0;
            if (section_infos != null)
            {
                results_cell_count += section_infos.Select(x => x.RowCount * x.Query.Columns.Count).Sum();
            }

            List<SectionQueryOutput<T>> sections = null;
            if (section_infos != null)
            {
               sections = new List<SectionQueryOutput<T>>(section_infos.Count);
                foreach (var section_info in section_infos)
                {
                    var section_output = new SectionQueryOutput<T>(section_info.RowCount, section_info.Query.SectionIndex);

                    int num_cols = section_info.Query.Columns.Count;
                    foreach (int row_index in section_info.RowIndexes)
                    {
                        var segment = segReader.GetNextSegment(num_cols);
                        var sec_res_row = new SectionQueryOutputRow<T>(segment, section_info.Query.SectionIndex, row_index);
                        section_output.Rows.Add(sec_res_row);
                    }

                    sections.Add(section_output);
                }
            }

            var output = new SectionsQueryOutput<T>(shapeid, results_cell_count, sections);
            
            int final_seg_size = segReader.Count;

            if ((final_seg_size - original_seg_size) != output.__totalcellcount)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Unexpected cursor");
            }

            return output;
        }

        private void cache_section_info(IList<Microsoft.Office.Interop.Visio.Shape> shapes)
        {
            // there aren't any subqueries so return an empty cache
            if (this.SectionQueries.Count < 1)
            {
                this._cache = new SectionInfoCache(0);
            }

            this._cache = new SectionInfoCache();

            // For each shape, for each section find the number of rows
            foreach (var shape in shapes)
            {
                var l_sectioninfo = new List<SectionInfo>(this.SectionQueries.Count);
                l_sectioninfo.AddRange(this.SectionQueries.Select(sec => sec.GetSectionInfoForShape(shape)));
                _cache.AddSectionInfosForShape(l_sectioninfo);
            }

            if (shapes.Count != _cache.CountShapes)
            {
                string msg = string.Format("mismatch in number of shapes and information collected for shapes");
                throw new VisioAutomation.Exceptions.InternalAssertionException(msg);
            }
        }

        private int _get_total_cell_count(int numshapes)
        {
            // Count the cells not in sections
            int count = 0;

            // Count the Cells in the Sections
            foreach (var section_info in this._cache.EnumSectionInfoForShapes)
            {
                count += section_info.Sum(s => s.RowCount * s.Query.Columns.Count);
            }

            return count;
        }

        private Streams.StreamArray _build_src_stream()
        {
            int dummy_shapeid = -1;
            int numshapes = 1;
            int shapeindex = 0;
            int numcells = this._get_total_cell_count(numshapes);
            var stream = new VisioAutomation.ShapeSheet.Streams.FixedSrcStreamBuilder(numcells);
            var cellinfos = this._enum_total_cellinfo(dummy_shapeid, shapeindex);
            var srcs = cellinfos.Select(i => i.SidSrc.Src);
            stream.AddRange(srcs);

            return stream.ToStream();
        }

        private VisioAutomation.ShapeSheet.Streams.StreamArray _build_sidsrc_stream(IList<int> shapeids)
        {
            int numshapes = shapeids.Count;
            int numcells = this._get_total_cell_count(numshapes);

            var stream = new VisioAutomation.ShapeSheet.Streams.FixedSidSrcStreamBuilder(numcells);

            for (int shapeindex = 0; shapeindex < shapeids.Count; shapeindex++)
            {
                // For each shape add the cells to query
                var shapeid = shapeids[shapeindex];

                var cellinfos = this._enum_total_cellinfo(shapeid, shapeindex);
                var sidsrcs = cellinfos.Select(i => i.SidSrc);
                stream.AddRange(sidsrcs);
            }

            return stream.ToStream();
        }

        private IEnumerable<Internal.QueryCellInfo> _enum_total_cellinfo(int shapeid, int shapeindex)
        {
            // enum SubQueries
            if (this._cache.CountShapes > 0)
            {
                var section_infos = _cache.GetSectionInfosForShapeAtIndex(shapeindex);
                foreach (var section_info in section_infos)
                {
                    foreach (int rowindex in section_info.RowIndexes)
                    {
                        foreach (var col in section_info.Query.Columns)
                        {
                            var src = new VisioAutomation.ShapeSheet.Src(
                                (short)section_info.Query.SectionIndex,
                                (short)rowindex,
                                col.CellIndex);
                            var sidsrc = new VisioAutomation.ShapeSheet.SidSrc((short)shapeid, src);
                            var cellinfo = new Internal.QueryCellInfo(sidsrc, col);
                            yield return cellinfo;
                        }
                    }
                }
            }
        }
    }
}