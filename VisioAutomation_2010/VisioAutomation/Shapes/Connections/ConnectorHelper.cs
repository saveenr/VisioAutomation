using VisioAutomation.Extensions;
using VA=VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Connections
{
    public static class ConnectorHelper
    {
        public static void ConnectShapes(IVisio.Shape connector_shape, IVisio.Shape from_shape, IVisio.Shape to_shape)
        {
            if (connector_shape == null)
            {
                throw new System.ArgumentNullException("connector_shape");
            }

            if (from_shape == null)
            {
                throw new System.ArgumentNullException("from_shape");
            }

            if (to_shape == null)
            {
                throw new System.ArgumentNullException("to_shape");
            }

            if (connector_shape == from_shape)
            {
                throw new System.ArgumentException("connector cannot be the FROM shape");
            }

            if (connector_shape == to_shape)
            {
                throw new System.ArgumentException("connector cannot be the TO shape");
            }
            var src_beginx = VA.ShapeSheet.SRCConstants.BeginX;
            var src_endx = VA.ShapeSheet.SRCConstants.EndX;
            var connector_beginx = connector_shape.CellsSRC[src_beginx.Section, src_beginx.Row, src_beginx.Cell];
            var connector_endx = connector_shape.CellsSRC[src_endx.Section, src_endx.Row, src_endx.Cell];
            var from_cell = from_shape.CellsSRC[1, 1, 0];
            var to_cell = to_shape.CellsSRC[1, 1, 0];
            connector_beginx.GlueTo(from_cell);
            connector_endx.GlueTo(to_cell);
        }

        public static IList<IVisio.Shape> ConnectShapes(IVisio.Page page, IVisio.Master master, IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException("master");
            }

            if (fromshapes == null)
            {
                throw new System.ArgumentNullException("fromshapes");
            }

            if (toshapes == null)
            {
                throw new System.ArgumentNullException("toshapes");
            }

            if (fromshapes.Count != toshapes.Count)
            {
                throw new System.ArgumentException("must have same number of from and to shapes");
            }
            
            if (fromshapes.Count == 0)
            {
                return new List<IVisio.Shape>(0);
            }

            int num_connectors = fromshapes.Count;
            var connectors = new List<IVisio.Shape>(num_connectors);

            var masters = Enumerable.Repeat(master, num_connectors).ToList();
            var points = Enumerable.Range(0, num_connectors).Select(i => new VA.Drawing.Point(i*2.0, -2)).ToList();
            short [] con_shapeids = page.DropManyU(masters, points);
            var con_shapes = page.Shapes.GetShapesFromIDs(con_shapeids);

            for (int i = 0; i < num_connectors; i++)
            {
                var from_shape = fromshapes[i];
                var to_shape = toshapes[i];
                var connector = con_shapes[i];

                // Connect from Shape 1 to Shape2
                ConnectShapes(connector, from_shape, to_shape);

                connectors.Add(connector);
            }

            return connectors;
        }
    }
}