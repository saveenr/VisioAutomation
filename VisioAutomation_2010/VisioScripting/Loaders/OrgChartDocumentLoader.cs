
using VisioScripting.Extensions;
using VAORGCHART = VisioAutomation.Models.Documents.OrgCharts;


namespace VisioScripting.Builders;

public class OrgChartDocumentLoader
{
    public static VAORGCHART.OrgChartDocument LoadFromXml(Client client, string filename)
    {
        var xdoc = SXL.XDocument.Load(filename);
        return OrgChartDocumentLoader.LoadFromXml(client, xdoc);
    }

    public static VAORGCHART.OrgChartDocument LoadFromXml(Client client, SXL.XDocument xdoc)
    {
        var root = xdoc.Root;

        var dic = new Dictionary<string, VAORGCHART.Node>();
        VAORGCHART.Node ocroot = null;

        client.Output.WriteVerbose("Walking XML");

        foreach (var ev in root.Elements())
        {
            if (ev.Name == "shape")
            {
                string id = ev.Attribute("id").Value;
                string parentid = ev.GetAttributeValue("parentid", null);
                var name = ev.Attribute("name").Value;

                client.Output.WriteVerbose( "Loading shape: {0} {1} {2}", id, name, parentid);
                var new_ocnode = new VAORGCHART.Node(name);

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
        client.Output.WriteVerbose( "Finished Walking XML");
        var oc = new VAORGCHART.OrgChartDocument();
        oc.OrgCharts.Add(ocroot);
        client.Output.WriteVerbose( "Finished Creating OrgChart model");
        return oc;
    }
}