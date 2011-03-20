using System.Xml.Linq;
using VisioAutomation.VDX.Elements;
using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.Internal.Extensions;

namespace VisioAutomation.VDX
{
    public class VDXWriter
    {
        public VDXWriter()
        {
        }

        public void CreateVDX(Drawing vdoc, XDocument vdx_xml_doc)
        {
            if (vdoc == null)
            {
                throw new System.ArgumentNullException("vdoc");
            }

            if (vdx_xml_doc == null)
            {
                throw new System.ArgumentNullException("vdx_xml_doc");
            }

            _ModifyTemplate(vdx_xml_doc, vdoc);
        }

        public void CreateVDX(Drawing vdoc, XDocument vdx_xml_doc, string output_filename)
        {
            if (output_filename == null)
            {
                throw new System.ArgumentNullException("output_filename");
            }

            this.CreateVDX(vdoc,vdx_xml_doc);

            // important to use DisableFormatting - Visio is very sensitive to whitespace in the <Text> element when there is complex formatting
            var saveoptions = System.Xml.Linq.SaveOptions.DisableFormatting;
            vdx_xml_doc.Save(output_filename, saveoptions);
        }

        public static void CleanUpTemplate(XDocument vdx_xml_doc)
        {
            var root = vdx_xml_doc.Root;

            string ns_2003 = Constants.VisioXmlNamespace2003;

            // set document properties
            var docprops = root.ElementVisioSchema2003("DocumentProperties");
            docprops.RemoveElement(ns_2003 + "PreviewPicture");
            docprops.SetElementValue(ns_2003 + "Creator", "");
            docprops.SetElementValue(ns_2003 + "Company", "");

            // remove any pages
            var pages = root.ElementVisioSchema2003("Pages");
            pages.RemoveNodes();

            // Do not remove the FaceNames node - it contains fonts to which the template may be referring
            root.RemoveElement(ns_2003 + "Windows");
            root.RemoveElement(ns_2003 + "DocumentProperties");


            // TODO Add DocumentSettings
            var docsettings = root.ElementsVisioSchema2003("DocumentSettings");
            if (docsettings!=null)
            {
                docsettings.Remove();
            }
        }

        private void _ModifyTemplate(System.Xml.Linq.XDocument vdx_xml_doc, Elements.Drawing vdoc)
        {
            var root = vdx_xml_doc.Root;
            root.AddFirst(vdoc.DocumentProperties.ToXml());

            var xfacenames = root.ElementVisioSchema2003("FaceNames");
            xfacenames.RemoveAll();

            foreach (var vface in vdoc.Faces.Items)
            {
                vface.ToXml(xfacenames);
            }

            var xcolors = root.ElementVisioSchema2003("Colors");
            xcolors.RemoveAll();

            int ix = 0;
            foreach (var color in vdoc.Colors)
            {
                color.AddToElement(xcolors, ix++);
            }

            var xpages = root.ElementVisioSchema2003("Pages");

            foreach (var vpage in vdoc.Pages.Items)
            {
                vpage.AddToElement(xpages);
            }

            if (vdoc.Windows != null && vdoc.Windows.Count > 0)
            {
                var xwindows = XMLUtil.CreateVisioSchema2003Element("Windows");
                root.Add(xwindows);

                foreach (var window in vdoc.Windows)
                {
                    window.AddToElement(xwindows);
                }
            }
        }
    }
}