using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using VA = VisioAutomation;
using VAS = VisioAutomation.Scripting;
using IVisio= Microsoft.Office.Interop.Visio;
namespace VisioAutomation.Scripting.FlowChart
{
    public class FlowChartBuilder
    {
        public static IList<RenderItem> LoadFromXML(Session scriptingsession, string filename)
        {
            var xmldoc = XDocument.Load(filename);
            return LoadFromXML(scriptingsession, xmldoc);
        }

        public static IList<RenderItem> LoadFromXML(Session scriptingsession, XDocument xmldoc)
        {
            var renderitems = new List<RenderItem>();
            bool major_error = false;
            var page_els = xmldoc.Root.Elements("page");
            foreach (var page_el in page_els)
            {
                var node_ids = new HashSet<string>();
                var con_ids = new HashSet<string>();

                var renderer = new Layout.MSAGL.DirectedGraphLayout();
                var renderoptions_el = page_el.Element("renderoptions");
                GetRenderOptionsFromXml(renderoptions_el, renderer);

                var drawing = new Layout.MSAGL.Drawing();
                var shape_els = page_el.Element("shapes").Elements("shape");
                var con_els = page_el.Element("connectors").Elements("connector");

                var shape_infos = shape_els.Select(e => ShapeInfo.FromXml(scriptingsession, e)).ToList();
                var con_infos = con_els.Select(e => ConnectorInfo.FromXml(scriptingsession, e)).ToList();

                // ANALYZE 1

                scriptingsession.Write(VAS.OutputStream.Verbose,"Analyzing shape data...");
                foreach (var shape_info in shape_infos)
                {
                    scriptingsession.Write(VAS.OutputStream.Verbose,"shape {0}", shape_info.ID);

                    if (node_ids.Contains(shape_info.ID))
                    {
                        scriptingsession.Write(VAS.OutputStream.Verbose,"ERROR: Node \"{0}\" is already defined", shape_info.ID);
                        major_error = true;
                    }
                    else
                    {
                        node_ids.Add(shape_info.ID);
                    }
                }

                scriptingsession.Write(VAS.OutputStream.Verbose,"Analyzing connector data...");
                foreach (var con_info in con_infos)
                {
                    scriptingsession.Write(VAS.OutputStream.Verbose,"connector {0}", con_info.ID);

                    if (con_ids.Contains(con_info.ID))
                    {
                        scriptingsession.Write(VAS.OutputStream.Verbose,"ERROR: Connector \"{0}\" is already defined", con_info.ID);
                        major_error = true;
                    }
                    else
                    {
                        con_ids.Add(con_info.ID);
                    }

                    if (!node_ids.Contains(con_info.From))
                    {
                        scriptingsession.Write(VAS.OutputStream.Verbose,
                            "ERROR: Connector \"{0}\" references a nonexistent FROM Node \"{1}\"",
                            con_info.ID, con_info.From);
                        major_error = true;
                    }

                    if (!node_ids.Contains(con_info.To))
                    {
                        scriptingsession.Write(VAS.OutputStream.Verbose,
                            "ERROR: Connector \"{0}\" references a nonexistent TO Node \"{1}\"",
                            con_info.ID, con_info.To);
                        major_error = true;
                    }
                }

                if (major_error)
                {
                    scriptingsession.Write(VAS.OutputStream.Verbose,"Errors encountered in shape data. Stopping.");
                    System.Environment.Exit(-1);
                }

                scriptingsession.Write(VAS.OutputStream.Verbose,"Creating shape AutoLayout nodes");
                foreach (var shape_info in shape_infos)
                {
                    var al_shape = drawing.AddShape(shape_info.ID, shape_info.Label, shape_info.Stencil,
                                                    shape_info.Master);
                    al_shape.URL = shape_info.URL;
                    al_shape.CustomProperties = new Dictionary<string, VA.CustomProperties.CustomPropertyCells>();
                    foreach (var kv in shape_info.custprops)
                    {
                        al_shape.CustomProperties[kv.Key] = kv.Value;
                    }
                }

                scriptingsession.Write(VAS.OutputStream.Verbose,"Creating connector AutoLayout nodes");
                foreach (var con_info in con_infos)
                {
                    var def_connector_type = VA.Connections.ConnectorType.Curved;
                    var connectory_type = def_connector_type;

                    var from_shape = drawing.FindShape(con_info.From);
                    var to_shape = drawing.FindShape(con_info.To);

                    var def_con_color = new VA.Drawing.ColorRGB(0x000000);
                    var def_con_weight = 1.0 / 72.0;
                    var def_end_arrow = 2;
                    var al_connector = drawing.Connect(con_info.ID, from_shape, to_shape, con_info.Label,
                                                       connectory_type);

                    al_connector.ShapeCells = new VA.DOM.ShapeCells();
                    al_connector.ShapeCells.LineColor = VA.Convert.ColorToFormulaRGB(con_info.Element.AttributeAsColor("color", def_con_color));
                    al_connector.ShapeCells.LineWeight = con_info.Element.AttributeAsInches("weight", def_con_weight);
                    al_connector.ShapeCells.EndArrow = def_end_arrow;
                }

                scriptingsession.Write(VAS.OutputStream.Verbose,"Rendering AutoLayout...");

                var renderitem = new RenderItem(drawing, renderer);
                renderitems.Add(renderitem);
                scriptingsession.Write(VAS.OutputStream.Verbose,"Finished rendering AutoLayout");
            }

            return renderitems;
        }

        public static void RenderDiagrams(
            VA.Scripting.Session scriptingsession,
            IList<RenderItem> renderitems)
        {
            scriptingsession.Write(VAS.OutputStream.Verbose,"Start Rendering FlowChart");
            var app = scriptingsession.Application;


            if (renderitems.Count < 1)
            {
                return;
            }

            var doc = scriptingsession.Document.NewDocument();

            int num_expected_pages = renderitems.Count;

            var seqnum = new VA.Scripting.SequenceNumberGenerator(1);

            foreach (int i in Enumerable.Range(0, renderitems.Count))
            {
                var diagram = renderitems[i].Drawing;
                var renderer = renderitems[i].DirectedGraphLayout;

                scriptingsession.Write(VAS.OutputStream.Verbose,"Rendering page: {0}", seqnum.Next());

                var options = new Layout.MSAGL.LayoutOptions();
                options.UseDynamicConnectors = false;

                IVisio.Page page = null;
                if (doc.Pages.Count == 1)
                {
                    page = app.ActivePage;
                }
                else
                {
                    page = doc.Pages.Add();
                }

                diagram.Render(page, options);
                scriptingsession.Page.ResizeToFitContents(new VA.Drawing.Size(1.0, 1.0), true);
                scriptingsession.View.Zoom(VA.Scripting.Zoom.ToPage);

                scriptingsession.Write(VAS.OutputStream.Verbose,"Finished rendering page");
            }

            scriptingsession.Write(VAS.OutputStream.Verbose,"Finished rendering pages");
            scriptingsession.Write(VAS.OutputStream.Verbose,"Finished rendering flowchart.");
        }

        private static void GetRenderOptionsFromXml(XElement el, Layout.MSAGL.DirectedGraphLayout directed_graph_layout)
        {
            System.Func<string, bool> bool_converter = s => bool.Parse(s);
            System.Func<string, int> int_converter = s => int.Parse(s);
            System.Func<string, double> double_converter = (s) => double.Parse(s);

            directed_graph_layout.LayoutOptions.UseDynamicConnectors = VAS.XmlUtil.GetAttributeValue(el,"usedynamicconnectors", bool_converter);
            directed_graph_layout.LayoutOptions.ScalingFactor = VAS.XmlUtil.GetAttributeValue(el,"scalingfactor", double_converter);
        }
    }
}