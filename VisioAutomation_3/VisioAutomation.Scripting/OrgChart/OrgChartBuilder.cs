using System.Collections.Generic;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.OrgChart
{
    public class OrgChartBuilder
    {
        public static VA.Layout.OrgChart.Drawing LoadFromXML(Session scriptingsession, string filename)
        {
            var xdoc = System.Xml.Linq.XDocument.Load(filename);
            return LoadFromXML(scriptingsession, xdoc);
        }

        public static VA.Layout.OrgChart.Drawing LoadFromXML(Session scriptingsession,
                                                             System.Xml.Linq.XDocument xdoc)
        {
            var root = xdoc.Root;

            var dic = new Dictionary<string, VA.Layout.OrgChart.Node>();
            VA.Layout.OrgChart.Node ocroot = null;

            scriptingsession.Write(VA.Scripting.OutputStream.Verbose,"Walking XML");

            foreach (var ev in root.Elements())
            {
                if (ev.Name == "shape")
                {
                    string id = ev.Attribute("id").Value;
                    string parentid = VA.Scripting.XmlUtil.GetAttributeValue(ev, "parentid", null);
                    var name = ev.Attribute("name").Value;

                    scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Loading shape: {0} {1} {2}", id, name, parentid);
                    var new_ocnode = new VA.Layout.OrgChart.Node(name);

                    if (ocroot == null)
                    {
                        ocroot = new_ocnode;
                    }

                    dic[id] = new_ocnode;

                    if (parentid != null)
                    {
                        if (dic.ContainsKey(parentid))
                        {
                            var parent = dic[parentid];
                            parent.Children.Add(new_ocnode);
                        }
                    }
                }
            }
            scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Finished Walking XML");
            var oc = new VA.Layout.OrgChart.Drawing();
            oc.Root = ocroot;
            scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Finished Creating OrgChart model");
            return oc;
        }
    }
}