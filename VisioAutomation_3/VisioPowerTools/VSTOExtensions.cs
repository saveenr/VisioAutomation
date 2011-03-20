using System.Collections.Generic;
using System.Linq;

namespace VisioPowerTools
{
    public static class VSTOExtensions
    {
        private static readonly System.Object missing = System.Type.Missing;

        public static Microsoft.Office.Core.CommandBarPopup AddNewPopup(this Microsoft.Office.Core.CommandBarPopup cmdbarpopup, string tag, string caption)
        {
            var new_popup = (Microsoft.Office.Core.CommandBarPopup)cmdbarpopup.Controls.Add(
                                                 Microsoft.Office.Core.MsoControlType.msoControlPopup, // Type
                                                 missing, // Object
                                                 missing, // Id
                                                 cmdbarpopup.Controls.Count + 1, // Before
                                                 true // Temporary 
                                                 );

            new_popup.Caption = caption;
            new_popup.Tag = tag;
            return new_popup;
        }

        public static Microsoft.Office.Core.CommandBarButton AddNewMenuItem(this Microsoft.Office.Core.CommandBarPopup cmdbarpopup,
                                                       string tag,
                                                       string caption,
                                                       Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler handler)
        {
            var before_pos = cmdbarpopup.Controls.Count + 1;

            var btn = (Microsoft.Office.Core.CommandBarButton)cmdbarpopup.Controls.Add(
                                            Microsoft.Office.Core.MsoControlType.msoControlButton, // Type
                                            missing, // Object
                                            missing, // Id
                                            before_pos, // Before
                                            true // Temporary
                                            );

            btn.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;
            btn.Caption = caption;
            btn.Tag = tag;
            btn.Click += handler;

            return btn;
        }


        public static IEnumerable<Microsoft.Office.Core.CommandBarControl> EnumControls(this Microsoft.Office.Core.CommandBarControls controls)
        {
            foreach (int i in Enumerable.Range(0, controls.Count))
            {
                var item = controls[i + 1];
                yield return item;
            }
        }

    }
}