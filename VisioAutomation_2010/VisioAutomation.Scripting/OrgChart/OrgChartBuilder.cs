using System.Collections.Generic;
using VA = VisioAutomation;
using OCMODEL = VisioAutomation.Layout.Models.OrgChart;

namespace VisioAutomation.Scripting.OrgChart
{
    public class OrgChartBuilder
    {
        public static OCMODEL.Drawing LoadFromXML(Session scriptingsession, string filename)
        {
            var xdoc = System.Xml.Linq.XDocument.Load(filename);
            return LoadFromXML(scriptingsession, xdoc);
        }

        public static OCMODEL.Drawing LoadFromXML(Session scriptingsession,
                                                             System.Xml.Linq.XDocument xdoc)
        {
            var root = xdoc.Root;

            var dic = new Dictionary<string, OCMODEL.Node>();
            OCMODEL.Node ocroot = null;

            scriptingsession.Write(VA.Scripting.OutputStream.Verbose,"Walking XML");

            foreach (var ev in root.Elements())
            {
                if (ev.Name == "shape")
                {
                    string id = ev.Attribute("id").Value;
                    string parentid = VA.Scripting.XmlUtil.GetAttributeValue(ev, "parentid", null);
                    var name = ev.Attribute("name").Value;

                    scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Loading shape: {0} {1} {2}", id, name, parentid);
                    var new_ocnode = new OCMODEL.Node(name);

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
            var oc = new OCMODEL.Drawing();
            oc.Root = ocroot;
            scriptingsession.Write(VA.Scripting.OutputStream.Verbose, "Finished Creating OrgChart model");
            return oc;
        }
    }
}