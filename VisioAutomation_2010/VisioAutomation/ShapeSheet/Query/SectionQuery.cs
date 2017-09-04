using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQuery
    {
        public SubQueryCollection SubQueries { get; }

        private SectionInfoCache _cache;

        public SectionQuery()
        {
            this.SubQueries = new SubQueryCollection(0);
        }

        public SubQuery AddSubQuery(Microsoft.Office.Interop.Visio.VisSectionIndices section)
        {
            var col = this.SubQueries.Add(section);
            return col;
        }

        private static void RestrictToShapesOnly(SurfaceTarget surface)
        {
            if (surface.Shape == null)
            {
                string msg = "Target must be Shape not Page or Master";
                throw new System.ArgumentException(msg);
            }
        }

        public QueryOutputSections<string> GetFormulas(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetFormulas(surface);
        }

        public QueryOutputSections<string> GetFormulas(SurfaceTarget surface)
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

        public QueryOutputSections<TResult> GetResults<TResult>(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return GetResults<TResult>(surface);
        }

        public QueryOutputSections<TResult> GetResults<TResult>(SurfaceTarget surface)
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

        public QueryOutputSections<ShapeSheet.CellData> GetFormulasAndResults(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var surface = new SurfaceTarget(shape);
            return this.GetFormulasAndResults(surface);
        }

        private static CellData[] _combine_formulas_and_results(string[] formulas, string[] results)
        {
            int n = results.Length;

            if (formulas.Length != results.Length)
            {
                throw new System.ArgumentException("Array Lengths must match");
            }

            var combined_data = new ShapeSheet.CellData[n];
            for (int i = 0; i < n; i++)
            {
                combined_data[i] = new ShapeSheet.CellData(formulas[i], results[i]);
            }
            return combined_data;
        }

        public QueryOutputSections<ShapeSheet.CellData> GetFormulasAndResults(SurfaceTarget surface)
        {
            RestrictToShapesOnly(surface);

            var shapes = new List<Microsoft.Office.Interop.Visio.Shape> { surface.Shape };

            this.cache_section_info(shapes);
            var srcstream = this._build_src_stream();
            const object[] unitcodes = null;
            var formulas = surface.GetFormulasU(srcstream);
            var results = surface.GetResults<string>(srcstream, unitcodes);
            var combined_data = _combine_formulas_and_results(formulas, results);

            var shape_index = 0;
            var sectioninfo = this.GetSectionInfoForShape(shape_index, _cache);
            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<CellData>(combined_data);
            var output_for_shape = this._create_output_for_shape(surface.ID16, sectioninfo, seg_builder);
            return output_for_shape;
        }

        public QueryOutputCollectionSections<string> GetFormulas(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetFormulas(surface, shapeids);
        }

        public QueryOutputCollectionSections<string> GetFormulas(SurfaceTarget surface, IList<int> shapeids)
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

        public QueryOutputCollectionSections<TResult> GetResults<TResult>(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetResults<TResult>(surface, shapeids);
        }

        public QueryOutputCollectionSections<TResult> GetResults<TResult>(SurfaceTarget surface, IList<int> shapeids)
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

        public QueryOutputCollectionSections<ShapeSheet.CellData> GetFormulasAndResults(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var surface = new SurfaceTarget(page);
            return this.GetFormulasAndResults(surface, shapeids);
        }

        public QueryOutputCollectionSections<ShapeSheet.CellData> GetFormulasAndResults(SurfaceTarget surface, IList<int> shapeids)
        {
            var shapes = new List<Microsoft.Office.Interop.Visio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Shapes.ItemFromID16[(short)shapeid]));

            this.cache_section_info(shapes);
            var srcstream = this._build_sidsrc_stream(shapeids);
            const object[] unitcodes = null;
            var results = surface.GetResults<string>(srcstream, unitcodes);
            var formulas = surface.GetFormulasU(srcstream);
            var combined_data = _combine_formulas_and_results(formulas, results);

            var seg_builder = new VisioAutomation.Utilities.ArraySegmentReader<CellData>(combined_data);
            var r = this._create_outputs_for_shapes(shapeids, _cache, seg_builder);
            return r;
        }

        private QueryOutputCollectionSections<T> _create_outputs_for_shapes<T>(IList<int> shapeids, SectionInfoCache cache, VisioAutomation.Utilities.ArraySegmentReader<T> segReader)
        {
            var output_for_all_shapes = new QueryOutputCollectionSections<T>();

            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shapeid = shapeids[shape_index];
                var subqueryinfo = this.GetSectionInfoForShape(shape_index, cache);
                var output_for_shape = this._create_output_for_shape((short)shapeid, subqueryinfo, segReader);
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

        private QueryOutputSections<T> _create_output_for_shape<T>(short shapeid, List<SectionInfo> section_infos, VisioAutomation.Utilities.ArraySegmentReader<T> segReader)
        {
            int original_seg_size = segReader.Count;

            int results_cell_count = 0;
            if (section_infos != null)
            {
                results_cell_count += section_infos.Select(x => x.RowCount * x.SubQuery.Columns.Count).Sum();
            }

            List<SubQueryOutput<T>> sections = null;
            if (section_infos != null)
            {
               sections = new List<SubQueryOutput<T>>(section_infos.Count);
                foreach (var section_info in section_infos)
                {
                    var subquery_output = new SubQueryOutput<T>(section_info.RowCount, section_info.SubQuery.SectionIndex);

                    int num_cols = section_info.SubQuery.Columns.Count;
                    foreach (int row_index in section_info.RowIndexes)
                    {
                        var segment = segReader.GetNextSegment(num_cols);
                        var sec_res_row = new SubQueryOutputRow<T>(segment, section_info.SubQuery.SectionIndex, row_index);
                        subquery_output.Rows.Add(sec_res_row);
                    }

                    sections.Add(subquery_output);
                }
            }

            var output = new QueryOutputSections<T>(shapeid, results_cell_count, sections);
            
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
            if (this.SubQueries.Count < 1)
            {
                this._cache = new SectionInfoCache(0);
            }

            this._cache = new SectionInfoCache();

            // For each shape, for each subquery (section) find the number of rows
            foreach (var shape in shapes)
            {
                var l_sectioninfo = new List<SectionInfo>(this.SubQueries.Count);
                l_sectioninfo.AddRange(this.SubQueries.Select(subquery => subquery.GetSectionInfoForShape(shape)));
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
                count += section_info.Sum(s => s.RowCount * s.SubQuery.Columns.Count);
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
                        foreach (var col in section_info.SubQuery.Columns)
                        {
                            var src = new VisioAutomation.ShapeSheet.Src(
                                (short)section_info.SubQuery.SectionIndex,
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