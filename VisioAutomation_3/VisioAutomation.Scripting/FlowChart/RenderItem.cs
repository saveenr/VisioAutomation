using VA=VisioAutomation;
using VAL = VisioAutomation.Layout;
using VAD = VisioAutomation.DOM;

namespace VisioAutomation.Scripting.FlowChart
{
    public class RenderItem
    {
        private Layout.MSAGL.Drawing _drawing;
        private Layout.MSAGL.DirectedGraphLayout _directed_graph_layout;

        public RenderItem(Layout.MSAGL.Drawing d, Layout.MSAGL.DirectedGraphLayout r)
        {
            this._drawing = d;
            this._directed_graph_layout = r;
        }

        public Layout.MSAGL.Drawing Drawing
        {
            get { return _drawing; }
            set { _drawing = value; }
        }

        public Layout.MSAGL.DirectedGraphLayout DirectedGraphLayout
        {
            get { return _directed_graph_layout; }
            set { _directed_graph_layout = value; }
        }
    }
}