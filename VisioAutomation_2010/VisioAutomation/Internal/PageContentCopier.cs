using VA=VisioAutomation;

namespace VisioAutomation.Internal
{
    internal class PageContentCopier
    {
        // page copying requires the clipboard
        // it occurs in two steps
        // first: the source page must be the active page and it's contents are placed on the clipboard
        // second: the contents of the clipboard are then pasted onto the destination page
        
        private bool has_clipboard_contents = false;
        short copy_paste_flags = (short)Microsoft.Office.Interop.Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate;
        VA.Pages.PageCells pagecells;
        
        public PageContentCopier(Microsoft.Office.Interop.Visio.Page src_page)
        {
            if (src_page == null)
            {
                throw new System.ArgumentNullException("Source Page is null");
            }

            var app = src_page.Application;

            if (src_page != app.ActivePage)
            {
                throw new System.ArgumentException("Source page must be active page.", "src_page");
            }

            var src_page_shapes = src_page.Shapes;
            int num_src_shapes = src_page_shapes.Count;

            if (num_src_shapes > 0)
            {
                has_clipboard_contents = true;
                var active_window = app.ActiveWindow;
                active_window.SelectAll();
                var selection = active_window.Selection;
   

                selection.Copy(copy_paste_flags);
                active_window.DeselectAll();
            }

            var src_pagesheet = src_page.PageSheet;
            pagecells = VA.Pages.PageCells.GetCells(src_pagesheet);

        }

        public void ApplyTo(Microsoft.Office.Interop.Visio.Page dest_page)
        {
            if (dest_page == null)
            {
                throw new System.ArgumentNullException("Destination Page is null");
            }

            var dest_pagesheet = dest_page.PageSheet;
            var update = new VisioAutomation.ShapeSheet.Update();
            pagecells.Apply(update);
            update.Execute(dest_pagesheet);
            
            if (this.has_clipboard_contents)
            {
                dest_page.Paste(this.copy_paste_flags);                
            }

        }
    }
}