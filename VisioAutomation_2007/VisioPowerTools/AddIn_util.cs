using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using IVisio = Microsoft.Office.Interop.Visio;
using MOC  = Microsoft.Office.Core;
using System.Linq;
using VA = VisioAutomation;


namespace VisioPowerTools
{
    public partial class VisioPowerToolsAddIn
    {
        private const string helper_menu_item_caption = "Power Tools";
        private const string menu_bar_name = "Menu Bar";
        private readonly static string tagprefixmenuitem = typeof(VisioPowerToolsAddIn).Name + "MenuItem";
        private readonly static string helper_menu_item_tag = typeof(VisioPowerToolsAddIn).Name + "HelperMenuItem";

        private MOC.CommandBars vis_cmd_bars;
        private List<MOC.CommandBarButton> buttons;
        public AddInData addinprefs;

        private MOC.CommandBarPopup add_popup(MOC.CommandBarPopup cmdbarpopup, string tag, string caption)
        {
            tag = tagprefixmenuitem + tag;
            return cmdbarpopup.AddNewPopup(tag, caption);
        }

        private MOC.CommandBarButton add_menu_item(MOC.CommandBarPopup cmdbarpopup,
                                               string tag,
                                               string caption,
                                               MOC._CommandBarButtonEvents_ClickEventHandler
                                                   handler)
        {
            tag = tagprefixmenuitem + tag;
            var btn = cmdbarpopup.AddNewMenuItem(tag, caption, (MOC.CommandBarButton Ctrl2, ref bool CancelDefault2) =>
                                                               wrap_handler(handler, Ctrl2, ref CancelDefault2));
            buttons.Add(btn);
            return btn;
        }

        private MOC.CommandBarButton add_menu_item(MOC.CommandBarPopup cmdbarpopup,
                                               string tag,
                                               string caption,
                                               System.Action handler)
        {
            tag = tagprefixmenuitem + tag;
            var btn = cmdbarpopup.AddNewMenuItem(tag, caption, (MOC.CommandBarButton Ctrl2, ref bool CancelDefault2) =>
                                                               wrap_handler(handler));
            buttons.Add(btn);
            return btn;
        }

        private void remove_duplicate_menu_items()
        {
            var cmdbar = vis_cmd_bars[menu_bar_name];
            var all_menu_items = cmdbar.Controls.EnumControls();

            var menu_items_to_remove = all_menu_items
                .Where(i => i.Tag == helper_menu_item_tag)
                .Cast<MOC.CommandBarPopup>()
                .ToList();

            foreach (var menu_item in menu_items_to_remove)
            {
                menu_item.Delete(null);
            }
        }

        private void wrap_handler(MOC._CommandBarButtonEvents_ClickEventHandler handler,
                                  MOC.CommandBarButton Ctrl,
                                  ref bool CancelDefault)
        {
            try
            {
                handler(Ctrl, ref CancelDefault);
            }
            catch (VA.AutomationException exc)
            {
                process_exception(exc, typeof (VA.AutomationException).Name);
            }
            catch (COMException exc)
            {
                process_exception(exc, typeof (COMException).Name);
            }
            catch (System.Exception exc)
            {
                process_exception(exc, typeof(System.Exception).Name);
                throw;
            }
        }

        private void wrap_handler(System.Action handler)
        {
            try
            {
                handler();
            }
            catch (VA.AutomationException exc)
            {
                process_exception(exc, typeof (VA.AutomationException).Name);
            }
            catch (COMException exc)
            {
                process_exception(exc, typeof (COMException).Name);
            }
            catch (System.Exception exc)
            {
                process_exception(exc, typeof(System.Exception).Name);
                throw;
            }
        }

        private static void process_exception(System.Exception exc, string prefix)
        {
            Debug.WriteLine(exc.Message);
            MessageBox.Show(prefix + "\n" + exc.Message);
        }
    }
}