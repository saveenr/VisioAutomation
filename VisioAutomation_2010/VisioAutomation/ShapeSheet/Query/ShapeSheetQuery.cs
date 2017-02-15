﻿using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeSheetQuery
    {
        public ColumnCellCollection Cells { get; }
        public SubQueryCollection SubQueries { get; }

        private SectionInfoCache cache;

        public ShapeSheetQuery()
        {
            this.Cells = new ColumnCellCollection(0);
            this.SubQueries = new SubQueryCollection(0);
        }

        public ColumnCell AddCell(ShapeSheet.SRC src, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            var col = this.Cells.Add(src, name);
            return col;
        }

        public SubQuery AddSubQuery(IVisio.VisSectionIndices section)
        {
            var col = this.SubQueries.Add(section);
            return col;
        }


        private static void RestrictToShapesOnly(ShapeSheetSurface surface)
        {
            if (surface.Target.Shape == null)
            {
                string msg = "Target must be Shape not Page or Master";
                throw new System.ArgumentException(msg);
            }
        }

        public QueryOutput<string> GetFormulas(ShapeSheetSurface surface)
        {
            RestrictToShapesOnly(surface);

            var shapes = new List<IVisio.Shape> { surface.Target.Shape };

            this.cache_section_info(shapes);
            var srcstream = this._build_src_stream();
            var values = surface.GetFormulasU(srcstream);
            var shape_index = 0;
            var cursor = 0;
            var sectioninfo = this.GetSectionInfoForShape(shape_index, cache);
            var output_for_shape = this._create_output_for_shape(surface.Target.ID16, values, sectioninfo, ref cursor);

            return output_for_shape;
        }

        public QueryOutput<TResult> GetResults<TResult>(ShapeSheetSurface surface)
        {
            RestrictToShapesOnly(surface);

            var shapes = new List<IVisio.Shape> { surface.Target.Shape };

            this.cache_section_info(shapes);
            var srcstream = this._build_src_stream();
            var unitcodes = this._build_unit_code_array(1);
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var shape_index = 0;
            var cursor = 0;
            var sectioninfo = this.GetSectionInfoForShape(shape_index, cache);
            var output_for_shape = this._create_output_for_shape(surface.Target.ID16, values, sectioninfo, ref cursor);
            return output_for_shape;
        }

        public QueryOutput<ShapeSheet.CellData> GetFormulasAndResults(ShapeSheetSurface surface)
        {
            RestrictToShapesOnly(surface);

            var shapes = new List<IVisio.Shape> { surface.Target.Shape };

            this.cache_section_info(shapes);
            var srcstream = this._build_src_stream();
            var unitcodes = this._build_unit_code_array(shapes.Count);
            var formulas = surface.GetFormulasU(srcstream);
            var results = surface.GetResults<string>(srcstream, unitcodes);
            var combined_data = CellData.Combine(formulas, results);

            var shape_index = 0;
            var cursor = 0;
            var sectioninfo = this.GetSectionInfoForShape(shape_index, cache);
            var output_for_shape = this._create_output_for_shape(surface.Target.ID16, combined_data, sectioninfo, ref cursor);
            return output_for_shape;
        }


        public QueryOutputCollection<string> GetFormulas(ShapeSheetSurface surface, IList<int> shapeids)
        {
            var shapes = new List<IVisio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Target.Shapes.ItemFromID16[(short)shapeid]));

            this.cache_section_info(shapes);
            var srcstream = this._build_sidsrc_stream(shapeids);
            var values = surface.GetFormulasU(srcstream);
            var list = this._create_outputs_for_shapes(shapeids, values, cache);
            return list;
        }

        public QueryOutputCollection<TResult> GetResults<TResult>(ShapeSheetSurface surface, IList<int> shapeids)
        {
            var shapes = new List<IVisio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Target.Shapes.ItemFromID16[(short)shapeid]));

            this.cache_section_info(shapes);
            var srcstream = this._build_sidsrc_stream(shapeids);
            var unitcodes = this._build_unit_code_array(shapeids.Count);
            var values = surface.GetResults<TResult>(srcstream, unitcodes);
            var list = this._create_outputs_for_shapes(shapeids, values, cache);
            return list;
        }

        public QueryOutputCollection<ShapeSheet.CellData> GetFormulasAndResults(ShapeSheetSurface surface, IList<int> shapeids)
        {
            var shapes = new List<IVisio.Shape>(shapeids.Count);
            shapes.AddRange(shapeids.Select(shapeid => surface.Target.Shapes.ItemFromID16[(short)shapeid]));

            this.cache_section_info(shapes);
            var srcstream = this._build_sidsrc_stream(shapeids);
            var unitcodes = this._build_unit_code_array(shapeids.Count);
            var results = surface.GetResults<string>(srcstream, unitcodes);
            var formulas  = surface.GetFormulasU(srcstream);
            var combined_data = CellData.Combine(formulas, results);
            var r = this._create_outputs_for_shapes(shapeids, combined_data, cache);
            return r;
        }

        private QueryOutputCollection<T> _create_outputs_for_shapes<T>(IList<int> shapeids, T[] values, SectionInfoCache cache)
        {
            var output_for_all_shapes = new QueryOutputCollection<T>();

            int cursor = 0;
            for (int shape_index = 0; shape_index < shapeids.Count; shape_index++)
            {
                var shapeid = shapeids[shape_index];
                var subqueryinfo = this.GetSectionInfoForShape(shape_index, cache);
                var output_for_shape =  this._create_output_for_shape((short)shapeid, values, subqueryinfo, ref cursor);
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

        private QueryOutput<T> _create_output_for_shape<T>(short shapeid, T[] values, List<SectionInfo> section_infos, ref int values_cursor)
        {
            int old_cursor = values_cursor;

            var output = new QueryOutput<T>(shapeid);

            // First Copy the Query Cell Values into the output
            output.Cells = new T[this.Cells.Count];
            for (int i = 0; i < this.Cells.Count; i++)
            {
                output.Cells[i] = values[values_cursor++];
            }

            // Now copy the Section values over
            if (section_infos != null)
            {
                output.Sections = new List<SubQueryOutput<T>>(section_infos.Count);
                foreach (var subquery_detail in section_infos)
                {
                    var subquery_output = new SubQueryOutput<T>(subquery_detail.RowCount);

                    int num_cols = subquery_detail.SubQuery.Columns.Count;
                    foreach (int row_index in subquery_detail.RowIndexes)
                    {
                        var row_values = new T[num_cols];
                        for (int col_index = 0; col_index < num_cols; col_index++)
                        {
                            row_values[col_index] = values[values_cursor++];
                        }
                        var sec_res_row = new SubQueryOutputRow<T>(row_values);
                        subquery_output.Rows.Add(sec_res_row);
                    }

                    output.Sections.Add(subquery_output);

                }
            }

            int num_cells = this.Cells.Count + ( section_infos == null ? 0 : section_infos.Select( x=>x.RowCount *x.SubQuery.Columns.Count).Sum());
            int expected_cursor = old_cursor + num_cells;
            if (expected_cursor != values_cursor)
            {
                throw new InternalAssertionException("Unexpected cursor");
            }

            return output;
        }

        private void cache_section_info(IList<IVisio.Shape> shapes)
        {
            // there aren't any subqueries so return an empty list
            if (this.SubQueries.Count < 1)
            {
                this.cache = new SectionInfoCache();
            }

            this.cache = new SectionInfoCache();

            // For each shape, for each subquery (section) find the number of rows
            foreach (var shape in shapes)
            {
                var l_sectioninfo = new List<SectionInfo>(this.SubQueries.Count);
                l_sectioninfo.AddRange(this.SubQueries.Select( subquery => subquery.GetSectionInfoForShape(shape)));
                cache.AddSectionInfosForShape(l_sectioninfo);
            }

            if (shapes.Count != cache.CountShapes)
            {
                string msg = string.Format("mismatch in number of shapes and information collected for shapes");
                throw new InternalAssertionException(msg);
            }
        }

        private int _get_total_cell_count(int numshapes)
        {
            // Count the cells not in sections
            int count = this.Cells.Count * numshapes;

            // Count the Cells in the Sections
            foreach (var data_for_shape in this.cache.EnumSectionInfoForShapes)
            {
                count += data_for_shape.Sum(s => s.RowCount*s.SubQuery.Columns.Count);
            }
            
            return count;
        }

        private short[] _build_src_stream()
        {
            int numshapes = 1;
            int shapeindex = 0;
            int numcells = this._get_total_cell_count(numshapes);
            var stream = new SRCStreamBuilder(numcells);

            int dummy_shapeid = -1;

            var qs = this.enum_cellinfo(dummy_shapeid, shapeindex);
            stream.AddRange(qs.Select(i => i.SIDSRC.SRC));


            if (stream.Count() != numcells)
            {
                string msg = string.Format("src list does not match expected size");
                throw new InternalAssertionException(msg);
            }

            return stream.ToStream();
        }

        private short[] _build_sidsrc_stream(IList<int> shapeids)
        {
            int numshapes = shapeids.Count;
            int numcells = this._get_total_cell_count(numshapes);

            var stream = new SIDSRCStreamBuilder(numcells);

            for (int shapeindex = 0; shapeindex < shapeids.Count; shapeindex++)
            {
                // For each shape add the cells to query
                var shapeid = shapeids[shapeindex];

                var qs = this.enum_cellinfo(shapeid, shapeindex);
                stream.AddRange(qs.Select(i => i.SIDSRC));
            }


            return stream.ToStream();
        }

        private IEnumerable<Internal.QueryCellInfo> enum_cellinfo(int shapeid, int shapeindex)
        {
            // enum Cells
            foreach (var col in this.Cells)
            {
                var sidsrc = new SIDSRC((short)shapeid, col.SRC);

                var q = new Internal.QueryCellInfo(sidsrc,col);
                yield return q;
            }

            // enum SubQueries
            if (this.cache.CountShapes > 0)
            {
                var section_infos = cache.GetSectionInfosForShapeAtIndex(shapeindex);
                foreach (var section_info in section_infos)
                {
                    foreach (int rowindex in section_info.RowIndexes)
                    {
                        foreach (var col in section_info.SubQuery.Columns)
                        {
                            var src = new VisioAutomation.ShapeSheet.SRC(
                                (short)section_info.SubQuery.SectionIndex,
                                (short)rowindex,
                                col.CellIndex);
                            var sidsrc = new VisioAutomation.ShapeSheet.SIDSRC((short)shapeid, src);
                            var q = new Internal.QueryCellInfo(sidsrc,col);
                            yield return q;
                        }
                    }
                }
            }
        }

        private UnitCodesBuilder _build_unit_code_array(int numshapes)
        {
            if (numshapes < 1)
            {
                throw new InternalAssertionException("numshapes must be >=1");
            }

            int numcells = this._get_total_cell_count(numshapes);

            
            var unitcodes = new UnitCodesBuilder(numcells);
            for (int shapeindex = 0; shapeindex < numshapes; shapeindex++)
            {
                // shapeindex - we aren't going to use it here so we don't care
                var infos = this.enum_cellinfo(-1, shapeindex);
                unitcodes.AddRange( infos.Select(i=>i.Column.UnitCode));
            }

            if (numcells != unitcodes.Count)
            {
                throw new InternalAssertionException("Number of unit codes must match number of cells");
            }

            return unitcodes;
        }
    }
}