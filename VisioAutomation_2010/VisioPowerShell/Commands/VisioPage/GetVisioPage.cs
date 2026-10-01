using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    // Parameter sets:
    //   "active"      -> -ActivePage switch:    return the active page.
    //   "pagebyid"    -> -ID <int[]>:           return pages with those Visio page IDs (Page.ID, so the
    //                                           first page of a new document is 0).
    //   "pagebyindex" -> -Index <int[]>:        return the pages at those 1-based positions in the document
    //                                           (Page.Index, so the first page is 1).
    //   "pagebyname"  -> -Name <string[]>:      return pages with those names.
    //                                           Also the DEFAULT set: a no-args call lands
    //                                           here with Name == null and returns every
    //                                           page on the (resolved) document.
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioPage, DefaultParameterSetName = "pagebyname")]
    public class GetVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, ParameterSetName = "active")]
        public SMA.SwitchParameter ActivePage;

        [SMA.Parameter(Position = 0, Mandatory = false, ParameterSetName = "pagebyname")]
        public string[] Name;

        [SMA.Parameter(Position = 0, Mandatory = false, ParameterSetName = "pagebyid")]
        public int[] ID;
        [SMA.Parameter(Mandatory = false, ParameterSetName = "pagebyindex")]
        public int[] Index;

        // CONTEXT:DOCUMENT
        [SMA.Parameter(Position = 1, Mandatory = false)]
        public IVisio.Document Document;
        
        protected override void ProcessRecord()
        {
            if (this.ActivePage)
            {
                var page_active = this.Client.Page.GetActivePage();
                this.WriteObject(page_active);
                return;
            }

            // If the active page  is not specified then work on all the pages in a document (user-specified or auto)

            var targetdoc = new VisioScripting.TargetDocument(this.Document);

            // First, the ID case: the real Visio page ID, as Get-VisioShape -ID does for shapes
            if (this.ID != null)
            {
                var t = targetdoc.ResolveToDocument(this.Client);
                var pages = t.Document.Pages;
                foreach (var id in this.ID)
                {
                    var page = pages.ItemFromID[id];
                    this.WriteObject(page);
                }
                return;
            }

            // Then the position case: Pages[n] is a 1-based index
            if (this.Index != null)
            {
                var t = targetdoc.ResolveToDocument(this.Client);
                var pages = t.Document.Pages;
                foreach (var index in this.Index)
                {
                    var page = pages[index];
                    this.WriteObject(page);
                }
                return;
            }

            // Then, handle the name case

            if (this.Name == null)
            {
                var pages_by_name = this.Client.Page.FindPagesInDocument(targetdoc, null);
                this.WriteObject(pages_by_name, true);
                return;
            }

            var list_page = new List<IVisio.Page>();
            foreach (var name in this.Name)
            {
                var pages_by_name = this.Client.Page.FindPagesInDocument(targetdoc, name);
                list_page.AddRange(pages_by_name);
            }
            this.WriteObject(list_page, true);
        }
    }
}