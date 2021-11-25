using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public static class ConnectorHelper
    {
        public static void ConnectShapes(IVisio.Shape from_shape, IVisio.Shape to_shape, IVisio.Shape connector_shape)
        {
            ConnectorHelper.ConnectShapes(from_shape, to_shape, connector_shape, true);
        }

        public static void ConnectShapes(IVisio.Shape from_shape, IVisio.Shape to_shape, IVisio.Shape connector_shape, bool manual_connection)
        {
            if (from_shape == null)
            {
                throw new System.ArgumentNullException(nameof(from_shape));
            }

            if (to_shape == null)
            {
                throw new System.ArgumentNullException(nameof(to_shape));
            }

            if (connector_shape == from_shape)
            {
                throw new System.ArgumentException("connector cannot be the FROM shape");
            }

            if (connector_shape == to_shape)
            {
                throw new System.ArgumentException("connector cannot be the TO shape");
            }

            if (manual_connection)
            {
                // Manuall Set the cells
                if (connector_shape == null)
                {
                    throw new System.ArgumentException("connector cannot be null when specifying manual connection");                    
                }

                var src_beginx = ShapeSheet.SrcConstants.OneDBeginX;
                var src_endx = ShapeSheet.SrcConstants.OneDEndX;
                var connector_beginx = connector_shape.CellsSRC[src_beginx.Section, src_beginx.Row, src_beginx.Cell];
                var connector_endx = connector_shape.CellsSRC[src_endx.Section, src_endx.Row, src_endx.Cell];
                var from_cell = from_shape.CellsSRC[1, 1, 0];
                var to_cell = to_shape.CellsSRC[1, 1, 0];
                connector_beginx.GlueTo(from_cell);
                connector_endx.GlueTo(to_cell);                                
            }
            else
            {
                // Use the AutoConnect feature
                if (connector_shape == null)
                {
                    from_shape.AutoConnect(to_shape, IVisio.VisAutoConnectDir.visAutoConnectDirNone);                    
                }
                else
                {
                    from_shape.AutoConnect(to_shape, IVisio.VisAutoConnectDir.visAutoConnectDirNone,connector_shape);                    
                    
                }
            }
        }

        public static List<IVisio.Shape> ConnectShapes( IVisio.Page page, IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes,
            IVisio.Master connector_master)
        {
            return ConnectorHelper.ConnectShapes(page, fromshapes, toshapes, connector_master, true);
        }

        public static List<IVisio.Shape> ConnectShapes(IVisio.Page page, IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes, IVisio.Master connector_master, bool force_manual)
        {
            if (connector_master == null && force_manual )
            {
                throw new System.ArgumentException("if the connector object is null then force manual must be false");                
            }
            // no_connector + force_manual -> INVALID
            // no_connector + not_force_manual -> AutoConect
            // yes_connector + force_manual -> Manual Connection
            // object false  + not_force_manual-> Autoconnect

            if (fromshapes == null)
            {
                throw new System.ArgumentNullException(nameof(fromshapes));
            }

            if (toshapes == null)
            {
                throw new System.ArgumentNullException(nameof(toshapes));
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

            var points = Enumerable.Range(0, num_connectors).Select(i => new Geometry.Point(i*2.0, -2)).ToList();
            IList<IVisio.Shape> con_shapes = null;
            if (connector_master != null)
            {
                var masters = Enumerable.Repeat(connector_master, num_connectors).ToList();
                short[] con_shapeids = page.DropManyU(masters, points);
                con_shapes = VisioAutomation.Shapes.ShapeHelper.GetShapesFromIDs(page.Shapes,con_shapeids);                
            }
            else
            {
                short[] con_shapeids = Pages.PageHelper.DropManyAutoConnectors(page, points);
                con_shapes = VisioAutomation.Shapes.ShapeHelper.GetShapesFromIDs(page.Shapes,con_shapeids);
            }

            for (int i = 0; i < num_connectors; i++)
            {
                var from_shape = fromshapes[i];
                var to_shape = toshapes[i];
                var connector = con_shapes[i];

                ConnectorHelper.ConnectShapes(from_shape, to_shape, connector, true);

                connectors.Add(connector);
            }

            return connectors;
        }
    }
}