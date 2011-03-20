using VisioAutomation;
using MOC = Microsoft.Office.Core;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioPowerTools
{
    public partial class VisioPowerToolsAddIn
    {
        private void add_anchor_window(System.Windows.Forms.Form form)
        {
            var application = this.Application;
            var parent_window = application.ActiveWindow;
            if (parent_window == null)
            {
                return;
            }
            if (application.ActiveDocument == null)
            {
                return;
            }

            object window_states = IVisio.VisWindowStates.visWSFloating | IVisio.VisWindowStates.visWSVisible;
            object window_types = IVisio.VisWinTypes.visAnchorBarAddon;

            var displacement = new System.Drawing.Point(50, 100);
            var window_rect = new System.Drawing.Rectangle(displacement, form.Size);
            string window_caption = form.Text;

            var the_anchor_window = VA.UI.UserInterfaceHelper.AddAnchorWindow(parent_window,
                                                                                       window_caption,
                                                                                       window_states,
                                                                                       window_types,
                                                                                       window_rect);

            if (the_anchor_window != null)
            {
                VA.UI.UserInterfaceHelper.AttachWindowsForm(the_anchor_window, form);
                form.Refresh();
            }
        }

        private void customize_visio_menu()
        {
            remove_duplicate_menu_items();

            var before_pos = vis_cmd_bars[menu_bar_name].Controls.Count + 1;
            var helper_main_menu_item = (MOC.CommandBarPopup) vis_cmd_bars[menu_bar_name].Controls.Add(
                MOC.MsoControlType.msoControlPopup, // Type
                missing, // Object
                missing, // Id
                before_pos, // Before
                true // Temporary
                                                              );

            helper_main_menu_item.Caption = helper_menu_item_caption;
            helper_main_menu_item.Tag = helper_menu_item_tag;

            var tools_menu = add_popup(helper_main_menu_item, "popup_tools", "Tools");
            var shape_menu = add_popup(helper_main_menu_item, "popup_shapes", "Shapes");
            var page_menu = add_popup(helper_main_menu_item, "popup_page", "Pages");
            var drawing_menu = add_popup(helper_main_menu_item, "popup_drawing", "Drawings");
            var importexport_menu = add_popup(helper_main_menu_item, "popup_export", "Import / Export");
            var dev_menu = add_popup(helper_main_menu_item, "popup_dev", "Developer");

            // Tools Menu
            add_menu_item(tools_menu, "menu_tools_fill", "Fill", CmdShowFillTool);
            add_menu_item(tools_menu, "menu_tools_color", "Color", CmdShowColorTool);
            add_menu_item(tools_menu, "menu_tools_text", "Text", CmdShowTextTool);
            add_menu_item(tools_menu, "menu_tools_formatpainter", "Format Painter", CmdShowFormatPainterTool);
            add_menu_item(tools_menu, "menu_tools_arrange", "Arrange", CmdShowArrangeTool);
            add_menu_item(tools_menu, "menu_tools_stacking", "Stacking", CmdShowStackingTool);
            add_menu_item(tools_menu, "menu_tools_layout", "Layout", CmdShowLayoutTool);
            add_menu_item(tools_menu, "menu_tools_selection", "Selection", CmdShowSelectionTool);
            add_menu_item(tools_menu, "menu_tools_view", "View", CmdShowViewTool);

            // Shape Menu
            add_menu_item(shape_menu, "menu_shape_stripws", "Strip Leading and Trailing Whitepace", CmdStripWhitespace);
            add_menu_item(shape_menu, "menu_shape_lock", "Lock Shape", CmdShapeLock);
            add_menu_item(shape_menu, "menu_shape_unlock", "Unlock Shape", CmdShapeUnlock);
            var shape_snappos_menu = add_popup(shape_menu, "menu_shape_snap_pos", "Snap Position");
            add_menu_item(shape_snappos_menu, "menu_shape_snappos1", "Snap Position to 1 inch",
                          CmdShapeSnapPositionOneInch);
            add_menu_item(shape_snappos_menu, "menu_shape_snappos2", "Snap Position to 1/2 inch",
                          CmdShapeSnapPositionHalfInch);
            add_menu_item(shape_snappos_menu, "menu_shape_snappos4", "Snap Position to 1/4 inch",
                          CmdShapeSnapPositionQuarterInch);
            add_menu_item(shape_snappos_menu, "menu_shape_snappos8", "Snap Position to 1/8 inch",
                          CmdShapeSnapPositionEighthInch);
            add_menu_item(shape_snappos_menu, "menu_shape_snappos16", "Snap Position to 1/16 inch",
                          CmdShapeSnapPositionSixteenthInch);
            var shape_snapsize_menu = add_popup(shape_menu, "menu_shape_snap_size", "Snap Size");
            add_menu_item(shape_snapsize_menu, "menu_shape_snapsize1", "Snap Size to 1 inch", CmdShapeSnapSizeOneInch);
            add_menu_item(shape_snapsize_menu, "menu_shape_snapsize2", "Snap Size to 1/2 inch", CmdShapeSnapSizeHalfInch);
            add_menu_item(shape_snapsize_menu, "menu_shape_snapsize4", "Snap Size to 1/4 inch",
                          CmdShapeSnapSizeQuarterInch);
            add_menu_item(shape_snapsize_menu, "menu_shape_snapsize8", "Snap Size to 1/8 inch",
                          CmdShapeSnapPositionEighthInch);
            add_menu_item(shape_snapsize_menu, "menu_shape_snapsize16", "Snap Size to 1/16 inch",
                          CmdShapeSnapPositionSixteenthInch);

            // Page Menu
            add_menu_item(page_menu, "menu_page_stripws", "Resize To Fit Contents", CmdPageResizeToFit);
            add_menu_item(page_menu, "menu_page_duplicate", "Duplicate", CmdPageDuplicate);
            add_menu_item(page_menu, "menu_page_duplicate", "Reset Origin", CmdPageResetOrigin);

            // Export Menu
            add_menu_item(importexport_menu, "menu_export_selection_as_svgxhtml", "Export Selection as SVG + XHTML", CmdExportAsSVGXHTML);
            add_menu_item(importexport_menu, "menu_export_selection_as_xamll", "Export Selection as XAML", CmdExportAsXAML);
            add_menu_item(importexport_menu, "menu_import_flowchart_xml", "Import FlowChart XML", CmdPageImportFlowChartXML);

            // Drawings Menu
            add_menu_item(drawing_menu, "menu_drawings_closeall", "Close all drawings without saving",
                          CmdCloseDocumentsWithoutSaving);
            add_menu_item(drawing_menu, "menu_drawings_closeapp", "Close app and all drawings without saving",
              CmdCloseDocumentsAndAppWithoutSaving);

            // Dev Menu
            add_menu_item(dev_menu, "menu_dev_throw_exception", "Throw Exception", CmdThrowException);
        }
    }
}