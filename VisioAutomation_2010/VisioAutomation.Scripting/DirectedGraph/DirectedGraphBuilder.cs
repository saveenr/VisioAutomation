using System.Collections.Generic;
using System.Linq;
using SXL = System.Xml.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using DGMODEL = VisioAutomation.Layout.Models.DirectedGraph;

namespace VisioAutomation.Scripting.DirectedGraph
{
    public class DirectedGraphBuilder
    {
        private class BuilderError
        {
            public string Text;

            public BuilderError(string text)
            {
                this.Text = text;
            }

            public static BuilderError ConnectorAlreadyDefined(int pagenum, string id)
            {
                return new BuilderError(string.Format("Page {0} : Connector \"{1}\" is already defined", pagenum, id));
            }

            public static BuilderError NodeAlreadyDefined(int pagenum, string id)
            {
                return new BuilderError(string.Format("Page {0} : Node \"{1}\" is already defined", pagenum, id));
            }

            public static BuilderError InvalidFromNode(int pagenum, string conid, string fromid)
            {
                return
                    new BuilderError(
                        string.Format("Page {0} : Connector \"{1}\" references a nonexistent FROM Node \"{2}\"",
                                      pagenum, conid, fromid));
            }

            public static BuilderError InvalidToNode(int pagenum, string conid, string toid)
            {
                return
                    new BuilderError(
                        string.Format("Page {0} : Connector \"{1}\" references a nonexistent TO Node \"{2}\"",
                                      pagenum, conid, toid));
            }
        }

        public static IList<DGMODEL.Drawing> LoadFromXML(VA.Scripting.Session scriptingsession, string filename)
        {
            var xmldoc = SXL.XDocument.Load(filename);
            return LoadFromXML(scriptingsession, xmldoc);
        }

        private class PageData
        {
            public int PageNumber;
            public VA.Layout.Models.DirectedGraph.MSAGLRenderer Renderer;
            public DGMODEL.Drawing Drawing;
            public List<ShapeInfo> ShapeInfos;
            public List<ConnectorInfo> ConnectorInfos;
            public List<BuilderError> Errors;
        }

        private static List<PageData> LoadPageDataFromXML(VA.Scripting.Session scriptingsession, SXL.XDocument xmldoc)
        {
            var pagedatas = new List<PageData>();
            // LOAD and ANALYZE EACH PAGE

            int pagenum = 0;
            var page_els = xmldoc.Root.Elements("page");

            foreach (var page_el in page_els)
            {
                var node_ids = new HashSet<string>();
                var con_ids = new HashSet<string>();

                var pagedata = new PageData();
                pagedatas.Add(pagedata);
                pagedata.PageNumber = pagenum++;
                pagedata.Errors = new List<BuilderError>();
                pagedata.Renderer = new VA.Layout.Models.DirectedGraph.MSAGLRenderer();
                var renderoptions_el = page_el.Element("renderoptions");
                GetRenderOptionsFromXml(renderoptions_el, pagedata.Renderer);

                pagedata.Drawing = new DGMODEL.Drawing();
                var shape_els = page_el.Element("shapes").Elements("shape");
                var con_els = page_el.Element("connectors").Elements("connector");

                pagedata.ShapeInfos = shape_els.Select(e => ShapeInfo.FromXml(scriptingsession, e)).ToList();
                pagedata.ConnectorInfos = con_els.Select(e => ConnectorInfo.FromXml(scriptingsession, e)).ToList();

                scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Analyzing shape data for page {0}", pagenum);
                foreach (var shape_info in pagedata.ShapeInfos)
                {
                    scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "shape {0}", shape_info.ID);

                    if (node_ids.Contains(shape_info.ID))
                    {
                        pagedata.Errors.Add(BuilderError.NodeAlreadyDefined(pagenum, shape_info.ID));
                    }
                    else
                    {
                        node_ids.Add(shape_info.ID);
                    }
                }

                scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Analyzing connector data...");
                foreach (var con_info in pagedata.ConnectorInfos)
                {
                    scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "connector {0}", con_info.ID);

                    if (con_ids.Contains(con_info.ID))
                    {
                        pagedata.Errors.Add(BuilderError.ConnectorAlreadyDefined(pagenum, con_info.ID));
                    }
                    else
                    {
                        con_ids.Add(con_info.ID);
                    }

                    if (!node_ids.Contains(con_info.From))
                    {
                        pagedata.Errors.Add(BuilderError.InvalidFromNode(pagenum, con_info.ID, con_info.From));
                    }

                    if (!node_ids.Contains(con_info.To))
                    {
                        pagedata.Errors.Add(BuilderError.InvalidToNode(pagenum, con_info.ID, con_info.To));
                    }
                }
            }

            return pagedatas;
        }


        public static IList<DGMODEL.Drawing> LoadFromXML(VA.Scripting.Session scriptingsession, SXL.XDocument xmldoc)
        {
            var pagedatas = LoadPageDataFromXML(scriptingsession, xmldoc);

            // STOP IF ANY ERRORS
            int num_errors = pagedatas.Select(pagedata => pagedata.Errors.Count).Sum();
            if (num_errors > 1)
            {
                foreach (var pagedata in pagedatas)
                {
                    foreach (var error in pagedata.Errors)
                    {
                        scriptingsession.Write(VA.Scripting.OutputStream.Verbose, error.Text);
                    }
                    scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Errors encountered in shape data. Stopping.");
                }
            }

            // DRAW EACH PAGE
            foreach (var pagedata in pagedatas)
            {
                scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Creating shape AutoLayout nodes");
                foreach (var shape_info in pagedata.ShapeInfos)
                {
                    var al_shape = pagedata.Drawing.AddShape(shape_info.ID, shape_info.Label, shape_info.Stencil,
                                                             shape_info.Master);
                    al_shape.URL = shape_info.URL;
                    al_shape.CustomProperties = new Dictionary<string, VA.CustomProperties.CustomPropertyCells>();
                    foreach (var kv in shape_info.custprops)
                    {
                        al_shape.CustomProperties[kv.Key] = kv.Value;
                    }
                }

                scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Creating connector AutoLayout nodes");
                foreach (var con_info in pagedata.ConnectorInfos)
                {
                    var def_connector_type = VA.Connections.ConnectorType.Curved;
                    var connectory_type = def_connector_type;

                    var from_shape = pagedata.Drawing.Shapes.Find(con_info.From);
                    var to_shape = pagedata.Drawing.Shapes.Find(con_info.To);

                    var def_con_color = new VA.Drawing.ColorRGB(0x000000);
                    var def_con_weight = 1.0/72.0;
                    var def_end_arrow = 2;
                    var al_connector = pagedata.Drawing.Connect(con_info.ID, from_shape, to_shape, con_info.Label,
                                                                connectory_type);

                    al_connector.Cells = new VA.DOM.ShapeCells();
                    al_connector.Cells.LineColor =
                        VA.Convert.ColorToFormulaRGB(con_info.Element.AttributeAsColor("color", def_con_color));
                    al_connector.Cells.LineWeight = con_info.Element.AttributeAsInches("weight", def_con_weight);
                    al_connector.Cells.EndArrow = def_end_arrow;
                }

                scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Rendering AutoLayout...");
            }
            scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Finished rendering AutoLayout");

            var drawings = pagedatas.Select(pagedata => pagedata.Drawing).ToList();
            return drawings;
        }

        public static void RenderDiagrams(
            VA.Scripting.Session scriptingsession,
            IList<DGMODEL.Drawing> drawings)
        {
            scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Start Rendering FlowChart");
            var app = scriptingsession.VisioApplication;


            if (drawings.Count < 1)
            {
                return;
            }

            var doc = scriptingsession.Document.New();
            int num_pages_created = 0;
            var doc_pages = doc.Pages;

            foreach (int i in Enumerable.Range(0, drawings.Count))
            {
                var directed_graph_drawing = drawings[i];


                scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Rendering page: {0}", i + 1);

                var options = new DGMODEL.MSAGLLayoutOptions();
                options.UseDynamicConnectors = false;

                IVisio.Page page = null;

                if (num_pages_created == 0)
                {
                    // if this is the first page to drawe
                    // then reuse the initial empty page in the document
                    page = app.ActivePage;
                }
                else
                {
                    // otherwise, create a new page.
                    page = doc_pages.Add();
                }

                directed_graph_drawing.Render(page, options);

                scriptingsession.Page.ResizeToFitContents(new VA.Drawing.Size(1.0, 1.0), true);
                scriptingsession.View.Zoom(VA.Scripting.Zoom.ToPage);

                scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Finished rendering page");

                num_pages_created++;
            }
            scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Finished rendering pages");
            scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Finished rendering flowchart.");
        }

        private static void GetRenderOptionsFromXml(SXL.XElement el, VA.Layout.Models.DirectedGraph.MSAGLRenderer directed_graph_layout)
        {
            System.Func<string, bool> bool_converter = s => bool.Parse(s);
            System.Func<string, int> int_converter = s => int.Parse(s);
            System.Func<string, double> double_converter = (s) => double.Parse(s);

            directed_graph_layout.LayoutOptions.UseDynamicConnectors = VA.Scripting.XmlUtil.GetAttributeValue(el,
                                                                                                              "usedynamicconnectors",
                                                                                                              bool_converter);
            directed_graph_layout.LayoutOptions.ScalingFactor = VA.Scripting.XmlUtil.GetAttributeValue(el,
                                                                                                       "scalingfactor",
                                                                                                       double_converter);
        }
    }
}