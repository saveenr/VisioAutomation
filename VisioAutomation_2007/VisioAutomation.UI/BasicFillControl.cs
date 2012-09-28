using System.ComponentModel;
using System.Windows.Forms;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.UI
{

    public enum FillPattern
    {
        None = 0,
        Solid = IVisio.VisCellVals.visSolid,
        WideUpDiagonal = IVisio.VisCellVals.visWideUpDiagonal,
        WideCross = IVisio.VisCellVals.visWideCross,
        WideDiagonalCross = IVisio.VisCellVals.visWideDiagonalCross,
        WideDownDiagonal = IVisio.VisCellVals.visWideDownDiagonal,
        WideHorz = IVisio.VisCellVals.visWideHorz,
        WideVert = IVisio.VisCellVals.visWideVert,
        BackDotsMini = IVisio.VisCellVals.visBackDotsMini,
        HalfAndHalf = IVisio.VisCellVals.visHalfAndHalf,
        ForeDotsMini = IVisio.VisCellVals.visForeDotsMini,
        ForeDotsNarrow = IVisio.VisCellVals.visForeDotsNarrow,
        ForeDotsWide = IVisio.VisCellVals.visForeDotsWide,
        ThickHorz = IVisio.VisCellVals.visThickHorz,
        ThickVertical = IVisio.VisCellVals.visThickVertical,
        ThickDownDiagonal = IVisio.VisCellVals.visThickDownDiagonal,
        ThickUpDiagonal = IVisio.VisCellVals.visThickUpDiagonal,
        ThickDiagonalCross = IVisio.VisCellVals.visThickDiagonalCross,
        BackDotsWide = IVisio.VisCellVals.visBackDotsWide,
        ThinHorz = IVisio.VisCellVals.visThinHorz,
        ThinVert = IVisio.VisCellVals.visThinVert,
        ThinDownDiagonal = IVisio.VisCellVals.visThinDownDiagonal,
        ThinUpDiagonal = IVisio.VisCellVals.visThinUpDiagonal,
        ThinCross = IVisio.VisCellVals.visThinCross,
        ThinDiagonalCross = IVisio.VisCellVals.visThinDiagonalCross,
        LinearLeftToRight = 25,
        LinearVertical = 26,
        LinearRightToLeft = 27,
        LinearTopToBottom = 28,
        LinearHorizontal = 29,
        LinearBottomToTop = 30,
        RectangularUpperLeft = 31,
        RectangularUpperRight = 32,
        RectangularLowerLeft = 33,
        RectangularLowerRight = 34,
        RectangularCenter = 35,
        RadialUpperLeft = 36,
        RadialUpperRight = 37,
        RadialLowerLeft = 38,
        RadialLowerRight = 39,
        RadialCenter = 40
    }

    public partial class BasicFillControl : UserControl
    {
        public BasicFillControl()
        {
            InitializeComponent();

            this.comboBoxPattern.DataSource = System.Enum.GetValues(typeof(FillPattern));
        }

        [Browsable(true)]
        public System.Drawing.Color ForegroundColor
        {
            get { return this.colorPickerForeground.Color; }
            set { this.colorPickerForeground.Color = value; }
        }

        [Browsable(true)]
        public System.Drawing.Color BackgroundColor
        {
            get { return this.colorPickerBackground.Color; }
            set { this.colorPickerBackground.Color = value; }
        }

        [Browsable(true)]
        public int ForegroundTransparency
        {
            get { return this.ucTransparency1.TransparencyPercent; }
            set { this.ucTransparency1.TransparencyPercent = value; }
        }

        [Browsable(true)]
        public int BackgroundTransparency
        {
            get { return this.ucTransparency2.TransparencyPercent; }
            set { this.ucTransparency2.TransparencyPercent = value; }
        }

        [Browsable(true)]
        public FillPattern FillPattern
        {
            get { return (FillPattern)this.comboBoxPattern.SelectedValue; }
            set { this.comboBoxPattern.SelectedItem = value; }
        }

        private void comboBoxGradient_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            var v = (FillPattern)this.comboBoxPattern.SelectedValue;
        }

        private void toolStripMenuItem1_Click(object sender, System.EventArgs e)
        {
            MessageBox.Show("!");
        }

        private void linkLabelTools_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var form = new FormBasicFillTools();
            form.ForegroundColor = this.ForegroundColor;
            form.BackgroundColor= this.BackgroundColor;

            var results = form.ShowDialog();
            if (results == DialogResult.OK)
            {
                this.ForegroundColor = form.ForegroundColor;
                this.BackgroundColor = form.BackgroundColor;
            }

        }
    }
}