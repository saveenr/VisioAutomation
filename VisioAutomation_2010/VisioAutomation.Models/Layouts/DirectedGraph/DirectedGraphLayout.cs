﻿using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class DirectedGraphLayout
    {
        public readonly IDList<Shape> Shapes;
        public readonly IDList<Connector> Connectors;

        public DirectedGraphLayout()
        {
            this.Shapes = new IDList<Shape>();
            this.Connectors = new IDList<Connector>();
        }

        public Shape AddShape(string id, string label, string stencil_name, string master_name)
        {
            var s0 = new Shape(id);
            s0.Label = label;
            s0.StencilName = stencil_name;
            s0.MasterName = master_name;

            this.Shapes.Add(id, s0);
            return s0;
        }

        public Connector AddConnection(
            string id,
            Shape from,
            Shape to,
            string label,
            ConnectorType type)
        {
            var new_connector = new Connector(from, to);
            new_connector.ID = id;
            new_connector.Label = label;
            new_connector.ConnectorType = type;

            this.Connectors.Add(id, new_connector);
            return new_connector;
        }
    }
}