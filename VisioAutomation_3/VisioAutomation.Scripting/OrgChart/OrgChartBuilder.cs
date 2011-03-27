using System.Collections.Generic;
using VA = VisioAutomation;
using VAS = VisioAutomation.Scripting;
using SXL = System.Xml.Linq;
using VAD = VisioAutomation.DOM;
using VAL = VisioAutomation.Layout;

namespace VisioAutomation.Scripting
{
     static class XmlUtil
    {
        public static string GetAttributeValue(SXL.XElement el, SXL.XName name, string defval)
        {
            var attr = el.Attribute(name);
            if (attr == null)
            {
                return defval;
            }

            return attr.Value ?? defval;
        }

        public static T GetAttributeValue<T>(SXL.XElement el, SXL.XName name, System.Func<string, T> converter)
        {
            var a = el.Attribute(name);
            if (a == null)
            {
                string msg = string.Format(System.Globalization.CultureInfo.InvariantCulture, "Missing value for attribute \"{0}\"", name);
                throw new System.ArgumentException(msg);
            }
            string v = el.Attribute(name).Value;
            return converter(v);
        }

        public static T GetAttributeValue<T>(SXL.XElement el, SXL.XName name, T defval, System.Func<string, T> converter)
        {
            var a = el.Attribute(name);
            if (a == null)
            {
                return defval;
            }
            string v = el.Attribute(name).Value;
            return converter(v);
        }
    }

}
namespace VisioAutomation.Scripting.OrgChart
{
    public class OrgChartBuilder
    {
        public static VAL.OrgChart.Drawing LoadFromXML(Session scriptingsession, string filename)
        {
            var xdoc = System.Xml.Linq.XDocument.Load(filename);
            return LoadFromXML(scriptingsession, xdoc);
        }

        public static VAL.OrgChart.Drawing LoadFromXML(Session scriptingsession,
                                                             System.Xml.Linq.XDocument xdoc)
        {
            var root = xdoc.Root;

            var dic = new Dictionary<string, VAL.OrgChart.Node>();
            VAL.OrgChart.Node ocroot = null;

            scriptingsession.Write(VAS.OutputStream.Verbose,"Walking XML");

            foreach (var ev in root.Elements())
            {
                if (ev.Name == "shape")
                {
                    string id = ev.Attribute("id").Value;
                    string parentid = VAS.XmlUtil.GetAttributeValue(ev,"parentid", null);
                    var name = ev.Attribute("name").Value;

                    scriptingsession.Write(VAS.OutputStream.Verbose,"Loading shape: {0} {1} {2}", id, name, parentid);
                    var new_ocnode = new VAL.OrgChart.Node(name);

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
            scriptingsession.Write(VAS.OutputStream.Verbose,"Finished Walking XML");
            var oc = new VAL.OrgChart.Drawing();
            oc.Root = ocroot;
            scriptingsession.Write(VAS.OutputStream.Verbose,"Finished Creating OrgChart model");
            return oc;
        }

        public static void RenderDiagrams(VA.Scripting.Session scriptingsession,
                                          VAL.OrgChart.Drawing drawing)
        {
            scriptingsession.Write(VAS.OutputStream.Verbose,"Start OrgChart Rendering");
            var renderer = new VAL.OrgChart.OrgChartLayout();
            var application = scriptingsession.VisioApplication;
            drawing.Render(application);
            var active_page = application.ActivePage;
            active_page.ResizeToFitContents();
            scriptingsession.Write(VAS.OutputStream.Verbose,"Finished OrgChart Rendering");
        }
    }
}