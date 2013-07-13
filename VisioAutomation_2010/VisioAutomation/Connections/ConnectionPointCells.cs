using VA=VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VisioAutomation.Extensions;

namespace VisioAutomation.Connections
{
    public class ConnectionPointCells : VA.ShapeSheet.CellGroups.CellGroupMultiRowEx
    {
        public VA.ShapeSheet.CellData<double> X { get; set; }
        public VA.ShapeSheet.CellData<double> Y { get; set; }
        public VA.ShapeSheet.CellData<int> DirX { get; set; }
        public VA.ShapeSheet.CellData<int> DirY { get; set; }
        public VA.ShapeSheet.CellData<int> Type { get; set; }

        public override void ApplyFormulasForRow(ApplyFormula func, short row)
        {
            func(VA.ShapeSheet.SRCConstants.Connections_X.ForRow(row), this.X.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_Y.ForRow(row), this.Y.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_DirX.ForRow(row), this.DirX.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_DirY.ForRow(row), this.DirY.Formula);
            func(VA.ShapeSheet.SRCConstants.Connections_Type.ForRow(row), this.Type.Formula);
        }

        private static ConnectionPointCells get_cells_from_row(ConnectionPointQuery query, VA.ShapeSheet.CellData<double>[] row)
        {
            var cells = new ConnectionPointCells();
            cells.X = row[query.X];
            cells.Y = row[query.Y];
            cells.DirX = row[query.DirX].ToInt();
            cells.DirY = row[query.DirY].ToInt();
            cells.Type = row[query.Type].ToInt();

            return cells;
        }

        public static IList<List<ConnectionPointCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {

            var outer_list = new List<List<ConnectionPointCells>>();

            var query = get_query();

            var data_for_shapes = query.GetFormulasAndResults<double>(page, shapeids);

            foreach (var  data_for_shape in data_for_shapes)
            {
                var inner_list = new List<ConnectionPointCells>();
                outer_list.Add(inner_list);

                var sec = data_for_shape.SectionCells[0];
                foreach (var row in sec.Rows)
                {
                    var cells = get_cells_from_row(query, row);
                    inner_list.Add(cells);
                }

            }

            return outer_list;
        }

        public static IList<ConnectionPointCells> GetCells(IVisio.Shape shape)
        {

            var query = get_query();

            var data_for_shape = query.GetFormulasAndResults<double>(shape);

            var inner_list = new List<ConnectionPointCells>();

            var sec = data_for_shape.SectionCells[0];
            foreach (var row in sec.Rows)
            {
                var cells = get_cells_from_row(query, row);
                inner_list.Add(cells);
            }

            return inner_list;
        }

        private static ConnectionPointQuery m_query;
        private static ConnectionPointQuery get_query()
        {
            m_query =  m_query ?? new ConnectionPointQuery();
            return m_query;
        }

        class ConnectionPointQuery : VA.ShapeSheet.Query.QueryEx
        {
            public int DirX { get; set; }
            public int DirY { get; set; }
            public int Type { get; set; }
            public int X { get; set; }
            public int Y { get; set; }

            public ConnectionPointQuery()
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionConnectionPts);
                DirX = sec.AddCell(VA.ShapeSheet.SRCConstants.Connections_DirX, "DirX");
                DirY = sec.AddCell(VA.ShapeSheet.SRCConstants.Connections_DirY, "DirY");
                Type = sec.AddCell(VA.ShapeSheet.SRCConstants.Connections_Type, "Type");
                X = sec.AddCell(VA.ShapeSheet.SRCConstants.Connections_X, "X");
                Y = sec.AddCell(VA.ShapeSheet.SRCConstants.Connections_Y, "Y");
            }
        }
    }


}