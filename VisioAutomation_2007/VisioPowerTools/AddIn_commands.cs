using IVisio = Microsoft.Office.Interop.Visio;
using MOC = Microsoft.Office.Core;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using System.Linq;

namespace VisioPowerTools
{
    public partial class VisioPowerToolsAddIn
    {
        private void CmdThrowException(MOC.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var application = this.Application;
            var documents = application.Documents;
            var doc = documents[0];
            string n = doc.Name;
        }

        private void CmdShowFillTool(MOC.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var form = new FormFillDesigner();
            this.add_anchor_window(form);
        }

        private void CmdShowColorTool(MOC.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var form = new FormColorTool();
            this.add_anchor_window(form);
        }

        private void CmdShowArrangeTool(MOC.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var form = new FormArrangeTool();
            this.add_anchor_window(form);
        }

        private void CmdShowStackingTool(MOC.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var form = new FormStackingTool();
            this.add_anchor_window(form);
        }

        private void CmdShowLayoutTool(MOC.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var form = new FormLayoutTool();
            this.add_anchor_window(form);
        }

        private void CmdShowSelectionTool(MOC.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var form = new FormSelectionTool();
            this.add_anchor_window(form);
        }

        private void CmdShowViewTool(MOC.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var form = new FormViewTool();
            this.add_anchor_window(form);
        }
        private void CmdCloseDocumentsAndAppWithoutSaving()
        {
            this.CmdCloseDocumentsWithoutSaving();
            var client = VisioPowerToolsAddIn.Client;
            var app = client.VisioApplication;

            app.Quit(true);

        }


        private void CmdCloseDocumentsWithoutSaving()
        {
            var client = VisioPowerToolsAddIn.Client;
            var app = client.VisioApplication;
            var docs = app.Documents;
            if (docs.Count<1)
            {
                return;
            }

            var unsaved_docs = docs.AsEnumerable()
                .Where(doc => doc.Saved == false).ToList();


            if (unsaved_docs.Count>0)
            {
                // there are some unsaved docs - we ask if it is ok to force them to go away without saving
                var result = System.Windows.Forms.MessageBox.Show(
                    "Close all documents without saving?",
                    "Confirm close",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question,
                    System.Windows.Forms.MessageBoxDefaultButton.Button2);

                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    client.Document.CloseAllWithoutSaving();
                }
            }
            else
            {
                // all the docs are saved - go ahead and close everything
                client.Document.CloseAllWithoutSaving();
            }

        }

        public static void CmdResizeFitToContents()
        {
            VisioPowerToolsAddIn.Client.Page.ResizeToFitContents(new VA.Drawing.Size(0, 0), true);
        }

        public static void CmdAlignLeft()
        {
            VisioPowerToolsAddIn.Client.Arrange.Align(null, VA.Drawing.AlignmentHorizontal.Left);
        }

        public static void CmdAlignCenter()
        {
            VisioPowerToolsAddIn.Client.Arrange.Align(null, VA.Drawing.AlignmentHorizontal.Center);
        }

        public static void CmdAlignRight()
        {
            VisioPowerToolsAddIn.Client.Arrange.Align(null, VA.Drawing.AlignmentHorizontal.Right);
        }

        public static void CmdAlignTop()
        {
            VisioPowerToolsAddIn.Client.Arrange.Align(null, VA.Drawing.AlignmentVertical.Top);
        }

        public static void CmdAlignMiddle()
        {
            VisioPowerToolsAddIn.Client.Arrange.Align(null, VA.Drawing.AlignmentVertical.Center);
        }

        public static void CmdAlignBottom()
        {
            VisioPowerToolsAddIn.Client.Arrange.Align(null, VA.Drawing.AlignmentVertical.Bottom);
        }

        public static void CmdSnapPosition()
        {
            var addin = Globals.VisioPowerToolsAddIn;
            double d = addin.addinprefs.SnapUnit;
            VisioPowerToolsAddIn.Client.Arrange.SnapCorner(d, d, VA.Arrange.SnapCornerPosition.LowerLeft);
        }

        public static void CmdSnapSize()
        {
            var addin = Globals.VisioPowerToolsAddIn;
            double d = addin.addinprefs.SnapUnit;
            VisioPowerToolsAddIn.Client.Arrange.SnapSize(null, d, d);
        }

        public static void CmdCopySize()
        {
            VisioPowerToolsAddIn.Client.Format.CopySize();
        }

        public static void CmdPasteSize()
        {
            VisioPowerToolsAddIn.Client.Format.PasteSize(null, true, true);
        }

        public static void CmdPasteWidth()
        {
            VisioPowerToolsAddIn.Client.Format.PasteSize(null, true, false);
        }

        public static void CmdPasteHeight()
        {
            VisioPowerToolsAddIn.Client.Format.PasteSize(null, false, true);
        }

        public static void CmdSwitchCase()
        {
            VisioPowerToolsAddIn.Client.Text.ToogleCase(null);
        }

        public static void CmdSelectAll()
        {
            VisioPowerToolsAddIn.Client.Selection.All();
        }

        public static void CmdSelectNone()
        {
            VisioPowerToolsAddIn.Client.Selection.None();
        }

        public static void CmdInvertSelection()
        {
            VisioPowerToolsAddIn.Client.Selection.Invert();
        }

        public static void CmdDistribute(IVisio.VisDistributeTypes s, bool gtg)
        {
            var addin = Globals.VisioPowerToolsAddIn;
            if (!VisioPowerToolsAddIn.Client.Selection.HasShapes())
            {
                return;
            }

            var application = addin.Application;
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            selection.Distribute(s, gtg);
        }

        public static void CmdDistributeHorizontalSpacing()
        {
            CmdDistribute(IVisio.VisDistributeTypes.visDistHorzSpace, false);
        }

        public static void CmdDistributeHorizontalCenter()
        {
            CmdDistribute(IVisio.VisDistributeTypes.visDistHorzCenter, false);
        }

        public static void CmdDistributeVerticalSpacing()
        {
            CmdDistribute(IVisio.VisDistributeTypes.visDistVertSpace, false);
        }

        public static void CmdDistributeVerticalMiddle()
        {
            CmdDistribute(IVisio.VisDistributeTypes.visDistVertMiddle, false);
        }


        public static void CmdZoomOnSelection()
        {
            VisioPowerToolsAddIn.Client.View.Zoom(VA.Scripting.Zoom.ToSelection);
        }

        public static void CmdPageResizeToFit()
        {
            VisioPowerToolsAddIn.Client.Page.ResizeToFitContents(new VA.Drawing.Size(0, 0), true);
        }

        public static void CmdPageDuplicate()
        {
            VisioPowerToolsAddIn.Client.Page.Duplicate();
        }

        public static void CmdPageResetOrigin()
        {
            VisioPowerToolsAddIn.Client.Page.ResetOrigin(VisioPowerToolsAddIn.Client.Page.Get());
        }

        public static void CmdExportAsSVGXHTML()
        {
            if (!VisioPowerToolsAddIn.Client.HasActiveDocument)
            {
                System.Windows.Forms.MessageBox.Show("Open or create a new Drawing to export it.");
                return;
            }

            var form = new FormExportSelectionAsFormat(FormExportSelectionAsFormat.enumExportFormat.ExportSVGXHTML);
            form.ShowDialog();
        }

        public static void CmdExportAsXAML()
        {
            if (!VisioPowerToolsAddIn.Client.HasActiveDocument)
            {
                System.Windows.Forms.MessageBox.Show("Open or create a new Drawing to export it.");
                return;
            }

            var form = new FormExportSelectionAsFormat(FormExportSelectionAsFormat.enumExportFormat.ExportXAML);
            form.ShowDialog();
        }

        public static void CmdPageImportFlowChartXML()
        {
            var form = new FormImportFlowChartXML();
            form.ShowDialog();
        }

        private void CmdShowTextTool(MOC.CommandBarButton ctrl, ref bool cancel_default)
        {
            var form = new FormTextTool();
            this.add_anchor_window(form);
        }

        private void CmdShowFormatPainterTool(MOC.CommandBarButton ctrl, ref bool cancel_default)
        {
            var form = new FormFormatPainter();
            this.add_anchor_window(form);
        }

        private const double one_inch = 1.0;
        private const double half_inch = 1.0/2.0;
        private const double quarter_inch = 1.0/4.0;
        private const double eighth_inch = 1.0/8.0;
        private const double sixteenth_inch = 1.0/16.0;

        public static void CmdShapeSnapPositionOneInch()
        {
            VisioPowerToolsAddIn.Client.Arrange.SnapCorner(one_inch, one_inch, VA.Arrange.SnapCornerPosition.LowerLeft);
        }

        public static void CmdShapeSnapPositionHalfInch()
        {
            VisioPowerToolsAddIn.Client.Arrange.SnapCorner(half_inch, half_inch, VA.Arrange.SnapCornerPosition.LowerLeft);
        }

        public static void CmdShapeSnapPositionQuarterInch()
        {
            VisioPowerToolsAddIn.Client.Arrange.SnapCorner(quarter_inch, quarter_inch, VA.Arrange.SnapCornerPosition.LowerLeft);
        }

        public static void CmdShapeSnapPositionEighthInch()
        {
            VisioPowerToolsAddIn.Client.Arrange.SnapCorner(eighth_inch, eighth_inch, VA.Arrange.SnapCornerPosition.LowerLeft);
        }

        public static void CmdShapeSnapPositionSixteenthInch()
        {
            VisioPowerToolsAddIn.Client.Arrange.SnapCorner(1.0 / 16.0, 1.0 / 16.0, VA.Arrange.SnapCornerPosition.LowerLeft);
        }

        public static void CmdShapeSnapSizeOneInch()
        {
            VisioPowerToolsAddIn.Client.Arrange.SnapSize(null, one_inch, one_inch);
        }

        public static void CmdShapeSnapSizeHalfInch()
        {
            VisioPowerToolsAddIn.Client.Arrange.SnapSize(null, half_inch, half_inch);
        }

        public static void CmdShapeSnapSizeQuarterInch()
        {
            VisioPowerToolsAddIn.Client.Arrange.SnapSize(null, quarter_inch, quarter_inch);
        }

        public static void CmdShapeSnapSizeEighthInch()
        {
            VisioPowerToolsAddIn.Client.Arrange.SnapSize(null, eighth_inch, eighth_inch);
        }

        public static void CmdShapeSnapSizeSixteenthInch()
        {
            VisioPowerToolsAddIn.Client.Arrange.SnapSize(null, sixteenth_inch, sixteenth_inch);
        }
    }
}