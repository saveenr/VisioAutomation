using System;
using SXL=System.Xml.Linq;
using VA=VisioAutomation;
using VisioAutomation.VDX.Internal.Extensions;


namespace VisioAutomation.VDX
{
    public class VDXWriter
    {
        public VDXWriter()
        {
        }

        public void CreateVDX(VA.VDX.Elements.Drawing vdoc, SXL.XDocument vdx_xml_doc)
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

        public void CreateVDX(VA.VDX.Elements.Drawing vdoc, SXL.XDocument vdx_xml_doc, string output_filename)
        {
            if (output_filename == null)
            {
                throw new System.ArgumentNullException("output_filename");
            }

            // Validate that all Document windows refer to an existing page
            foreach (var window in vdoc.Windows)
            {
                if (window is VA.VDX.Elements.DocumentWindow)
                {
                    var docwind = (VA.VDX.Elements.DocumentWindow) window;
                    docwind.ValidatePage(vdoc);
                }
            }
            this.CreateVDX(vdoc,vdx_xml_doc);

            // important to use DisableFormatting - Visio is very sensitive to whitespace in the <Text> element when there is complex formatting
            var saveoptions = SXL.SaveOptions.DisableFormatting;
            vdx_xml_doc.Save(output_filename, saveoptions);
        }

        public static void CleanUpTemplate(SXL.XDocument vdx_xml_doc)
        {
            var root = vdx_xml_doc.Root;

            string ns_2003 = VA.VDX.Internal.Constants.VisioXmlNamespace2003;

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


            // TODO Add DocumentSettings to VDX
            var docsettings = root.ElementsVisioSchema2003("DocumentSettings");
            if (docsettings!=null)
            {
                SXL.Extensions.Remove(docsettings);
            }
        }

        private void _ModifyTemplate(SXL.XDocument vdx_xml_doc, Elements.Drawing doc_node)
        {
            var root = vdx_xml_doc.Root;
            root.AddFirst(doc_node.DocumentProperties.ToXml());

            var xfacenames = root.ElementVisioSchema2003("FaceNames");
            xfacenames.RemoveAll();

            foreach (var vface in doc_node.Faces.Items)
            {
                vface.ToXml(xfacenames);
            }

            var xcolors = root.ElementVisioSchema2003("Colors");
            xcolors.RemoveAll();

            int ix = 0;
            foreach (var color in doc_node.Colors)
            {
                color.AddToElement(xcolors, ix++);
            }

            var xpages = root.ElementVisioSchema2003("Pages");

            foreach (var page_node in doc_node.Pages.Items)
            {
                page_node.AddToElement(xpages);
            }

            if (doc_node.Windows != null && doc_node.Windows.Count > 0)
            {
                var xwindows = VA.VDX.Internal.XMLUtil.CreateVisioSchema2003Element("Windows");
                root.Add(xwindows);

                foreach (var window in doc_node.Windows)
                {
                    window.AddToElement(xwindows);
                }
            }
        }
    }
}