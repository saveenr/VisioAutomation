using SXL=System.Xml.Linq;
using VA=VisioAutomation;
using VisioAutomation.VDX.Internal.Extensions;


namespace VisioAutomation.VDX
{
    internal class VDXWriter
    {
        public VDXWriter()
        {
        }

        public void CreateVDX(VA.VDX.Elements.Drawing vdoc, SXL.XDocument dom)
        {
            if (vdoc == null)
            {
                throw new System.ArgumentNullException("vdoc");
            }

            if (dom == null)
            {
                throw new System.ArgumentNullException("dom");
            }

            _ModifyTemplate(dom, vdoc);
        }

        public void CreateVDX(VA.VDX.Elements.Drawing vdoc, SXL.XDocument dom, string output_filename)
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
            this.CreateVDX(vdoc, dom);

            // important to use DisableFormatting - Visio is very sensitive to whitespace in the <Text> element when there is complex formatting
            var saveoptions = SXL.SaveOptions.DisableFormatting;

            dom.Save(output_filename, saveoptions);
        }

        private void _ModifyTemplate( SXL.XDocument dom, Elements.Drawing doc_node)
        {
            if (dom.Root == null)
            {
                throw new System.ArgumentException("DOM must have a root node");
            }

            var root = dom.Root;
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