/* Title:       Display any usercontrol as a popup menu
 * Author:      Pascal GANAYE (Originally in VB.NET, converted to C# by John Johnson II)
 * Email:       pascalcp@ganaye.com
 * Environment: C#.NET 2003
 * Keywords:    Popup, Contextual, Menu, Tooltip
 * Level:       Beginner
 * Description: This class lets you show any UserControl in an XP style popup menu.

 * feb 12, 2005 - Modification by Stumpy842 alias Steven Stover
 *                Added line to prevent showing in taskbar 
 *                Changed 4 dockpadding into one dockpadding.all
 * 
 * apr 14, 2005 - Modifiaction by John Johnson II
 *                Added public property to control the Opacity of the popup.
 */

using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace VisioAutomation.UI
{
    public class Popup : Component
    {
        public interface IPopupUserControl
        {
            bool AcceptPopupClosing();
        }

        public enum ePlacement
        {
            Left = 1,
            Right = 2,
            Top = 4,
            Bottom = 8,
            TopLeft = ePlacement.Top | ePlacement.Left,
            TopRight = ePlacement.Top | ePlacement.Right,
            BottomLeft = ePlacement.Bottom | ePlacement.Left,
            BottomRight = ePlacement.Bottom | ePlacement.Right
        }

        private Control mParent;

        private PopupForm mForm;

        public Popup(Control UserControl, Control parent)
        {
            this.Opacity = 1;
            this.AnimationSpeed = 150;
            this.ShowShadow = true;
            this.BorderColor = Color.DarkGray;
            this.HorizontalPlacement = ePlacement.BottomLeft;
            this.Resizable = false;
            this.mParent = parent;
            this.UserControl = UserControl;
        }

        public void Show()
        {
            // I use a shared variable in PopupForm class level for this ShowShadow
            // because the CreateParams is called from within the form constructor 
            // and we need a way to inform the form if a shadow is nescessary or not
            PopupForm.mShowShadow = this.ShowShadow;
            if (this.mForm != null)
            {
                this.mForm.DoClose();
            }
            this.mForm = new PopupForm(this);
            this.OnDropDown(this.mParent, new System.EventArgs());
        }

        // This internal class is a borderless form used to show the popup
        private class PopupForm : Form
        {
            public static bool mShowShadow;
            private bool mClosing;
            private const int BORDER_MARGIN = 1;
            private readonly Timer mTimer;
            private Size mControlSize;
            private Size mWindowSize = new Size(0, 0);
            private Point mNormalPos;
            private Rectangle mCurrentBounds = new Rectangle(0, 0, 0, 0);
            private readonly Popup mPopup;
            private readonly ePlacement mPlacement;
            private readonly System.DateTimeOffset mTimerStarted;
            private double mProgress;
            private int mx, my;
            private bool mResizing;
            public readonly Panel mResizingPanel;
            private const int CS_DROPSHADOW = 0x20000;
            private static Image mBackgroundImage;

            public event System.EventHandler DropDown;
            public event System.EventHandler DropDownClosed;

            public PopupForm(Popup popup)
            {
                this.mPopup = popup;
                this.SetStyle(ControlStyles.ResizeRedraw, true);
                this.FormBorderStyle = FormBorderStyle.None;
                this.StartPosition = FormStartPosition.Manual;
                this.ShowInTaskbar = false;
                this.DockPadding.All = PopupForm.BORDER_MARGIN;
                this.mControlSize = this.mPopup.UserControl.Size;
                this.mPopup.UserControl.Dock = DockStyle.Fill;
                this.Controls.Add(this.mPopup.UserControl);
                this.mWindowSize.Width = this.mControlSize.Width + 2*PopupForm.BORDER_MARGIN;
                this.mWindowSize.Height = this.mControlSize.Height + 2*PopupForm.BORDER_MARGIN;
                this.Opacity = popup.Opacity;

                //These are here to suppress warnings.
                this.DropDown += this.DoNothing;
                this.DropDownClosed += this.DoNothing;

                Form parentForm = this.mPopup.mParent.FindForm();
                if (parentForm != null)
                {
                    parentForm.AddOwnedForm(this);
                }

                if (this.mPopup.Resizable)
                {
                    this.mResizingPanel = new Panel();
                    if (PopupForm.mBackgroundImage == null)
                    {
                        var resources = new System.Resources.ResourceManager(typeof (Popup));
                        PopupForm.mBackgroundImage = (Image) resources.GetObject("CornerPicture.Image");
                    }
                    this.mResizingPanel.BackgroundImage = PopupForm.mBackgroundImage;
                    this.mResizingPanel.Width = 12;
                    this.mResizingPanel.Height = 12;
                    this.mResizingPanel.BackColor = Color.Red;
                    this.mResizingPanel.Left = this.mPopup.UserControl.Width - 15;
                    this.mResizingPanel.Top = this.mPopup.UserControl.Height - 15;
                    this.mResizingPanel.Cursor = Cursors.SizeNWSE;
                    this.mResizingPanel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
                    this.mResizingPanel.Parent = this;
                    this.mResizingPanel.BringToFront();

                    this.mResizingPanel.MouseUp += this.mResizingPanel_MouseUp;
                    this.mResizingPanel.MouseDown += this.mResizingPanel_MouseDown;
                    this.mResizingPanel.MouseMove += this.mResizingPanel_MouseMove;
                }
                this.mPlacement = this.mPopup.HorizontalPlacement;

                // Try to place the popup at the asked location
                this.ReLocate();

                // Check if the form is out of the screen
                // And if yes try to adapt the placement
                Rectangle workingArea = Screen.FromControl(this.mPopup.mParent).WorkingArea;
                if (this.mNormalPos.X + this.Width > workingArea.Right)
                {
                    if ((this.mPlacement & ePlacement.Right) != 0)
                    {
                        this.mPlacement = (this.mPlacement & ~ePlacement.Right) | ePlacement.Left;
                    }
                }
                else
                {
                    if (this.mNormalPos.X < workingArea.Left)
                    {
                        if ((this.mPlacement & ePlacement.Left) != 0)
                        {
                            this.mPlacement = (this.mPlacement & ~ePlacement.Left) | ePlacement.Right;
                        }
                    }
                }

                if (this.mNormalPos.Y + this.Height > workingArea.Bottom)
                {
                    if ((this.mPlacement & ePlacement.Bottom) != 0)
                    {
                        this.mPlacement = (this.mPlacement & ~ePlacement.Bottom) | ePlacement.Top;
                    }
                }
                else
                {
                    if (this.mNormalPos.Y < workingArea.Top)
                    {
                        if ((this.mPlacement & ePlacement.Top) != 0)
                        {
                            this.mPlacement = (this.mPlacement & ~ePlacement.Top) | ePlacement.Bottom;
                        }
                    }
                }

                if (this.mPlacement != this.mPopup.HorizontalPlacement)
                {
                    this.ReLocate();
                }

                // Check if the form is still out of the screen
                // If yes just move it back into the screen without changing Placement
                if (this.mNormalPos.X + this.mWindowSize.Width > workingArea.Right)
                {
                    this.mNormalPos.X = workingArea.Right - this.mWindowSize.Width;
                }
                else
                {
                    if (this.mNormalPos.X < workingArea.Left)
                    {
                        this.mNormalPos.X = workingArea.Left;
                    }
                }

                if (this.mNormalPos.Y + this.mWindowSize.Height > workingArea.Bottom)
                {
                    this.mNormalPos.Y = workingArea.Bottom - this.mWindowSize.Height;
                }
                else
                {
                    if (this.mNormalPos.Y < workingArea.Top)
                    {
                        this.mNormalPos.Y = workingArea.Top;
                    }
                }

                // Initialize the animation
                this.mProgress = 0;
                if (this.mPopup.AnimationSpeed > 0)
                {
                    this.mTimer = new Timer();

                    // I always aim 25 images per seconds.. seems to be a good value
                    // it looks smooth enough on fast computers and do not drain slower one
                    this.mTimer.Interval = 1000/25;
                    this.mTimerStarted = System.DateTimeOffset.Now;
                    this.mTimer.Tick += this.Showing;
                    this.mTimer.Start();
                    this.Showing(null, null);
                }
                else
                {
                    this.SetFinalLocation();
                }

                this.Show();
                this.mPopup.OnDropDown(this.mPopup.mParent, new System.EventArgs());
            }

            public static bool DropShadowSupported()
            {
                var os = System.Environment.OSVersion;
                return ((os.Platform == System.PlatformID.Win32NT) && (os.Version.CompareTo(new System.Version(5, 1, 0, 0)) >= 0));
            }

            protected override CreateParams CreateParams
            {
                get
                {
                    CreateParams parameters = base.CreateParams;
                    if (PopupForm.mShowShadow && PopupForm.DropShadowSupported())
                    {
                        parameters.ClassStyle = parameters.ClassStyle | PopupForm.CS_DROPSHADOW;
                    }
                    return parameters;
                }
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    if (this.mTimer != null)
                    {
                        this.mTimer.Dispose();
                    }
                }
                base.Dispose(disposing);
            }

            private void ReLocate()
            {
                int rW = this.mWindowSize.Width, rH = this.mWindowSize.Height;

                this.mNormalPos = this.mPopup.mParent.PointToScreen(Point.Empty);
                switch (this.mPlacement)
                {
                    case ePlacement.Top:
                    case ePlacement.TopLeft:
                    case ePlacement.TopRight:
                        this.mNormalPos.Y -= rH;
                        break;
                    case ePlacement.Bottom:
                    case ePlacement.BottomLeft:
                    case ePlacement.BottomRight:
                        this.mNormalPos.Y += this.mPopup.mParent.Height;
                        break;
                    case ePlacement.Left:
                    case ePlacement.Right:
                        this.mNormalPos.Y += (this.mPopup.mParent.Height - rH)/2;
                        break;
                }

                switch (this.mPlacement)
                {
                    case ePlacement.Left:
                        this.mNormalPos.X -= rW;
                        break;
                    case ePlacement.TopRight:
                    case ePlacement.BottomRight:
                        break;
                    case ePlacement.Right:
                        this.mNormalPos.X += this.mPopup.mParent.Width;
                        break;
                    case ePlacement.TopLeft:
                    case ePlacement.BottomLeft:
                        this.mNormalPos.X += this.mPopup.mParent.Width - rW;
                        break;
                    case ePlacement.Top:
                    case ePlacement.Bottom:
                        this.mNormalPos.X += (this.mPopup.mParent.Width - rW)/2;
                        break;
                }
            }

            private void Showing(object sender, System.EventArgs e)
            {
                this.mProgress = System.DateTimeOffset.Now.Subtract(this.mTimerStarted).TotalMilliseconds /this.mPopup.AnimationSpeed;
                if (this.mProgress >= 1)
                {
                    this.mTimer.Stop();
                    this.mTimer.Tick -= this.Showing;
                    this.AnimateForm(1);
                }
                else
                {
                    this.AnimateForm(this.mProgress);
                }
            }

            protected override void OnDeactivate(System.EventArgs e)
            {
                base.OnDeactivate(e);

                if (this.mClosing == false)
                {
                    if (this.mPopup.UserControl is IPopupUserControl)
                    {
                        this.mClosing = ((IPopupUserControl) this.mPopup.UserControl).AcceptPopupClosing();
                    }
                    else
                    {
                        this.mClosing = true;
                    }

                    if (this.mClosing)
                    {
                        this.DoClose();
                    }
                }
            }

            protected override void OnPaintBackground(PaintEventArgs pevent)
            {
                pevent.Graphics.DrawRectangle(new Pen(this.mPopup.BorderColor), 0, 0, this.Width - 1, this.Height - 1);
            }

            private void SetFinalLocation()
            {
                this.mProgress = 1;
                this.AnimateForm(1);
                this.Invalidate();
            }

            private void AnimateForm(double Progress)
            {
                double x = 0, y = 0, w = 0, h = 0;

                if (Progress <= 0.1)
                {
                    Progress = 0.1;
                }

                switch (this.mPlacement)
                {
                    case ePlacement.Top:
                    case ePlacement.TopLeft:
                    case ePlacement.TopRight:
                        y = 1 - Progress;
                        h = Progress;
                        break;
                    case ePlacement.Bottom:
                    case ePlacement.BottomLeft:
                    case ePlacement.BottomRight:
                        y = 0;
                        h = Progress;
                        break;
                    case ePlacement.Left:
                    case ePlacement.Right:
                        y = 0;
                        h = 1;
                        break;
                }

                switch (this.mPlacement)
                {
                    case ePlacement.TopRight:
                    case ePlacement.BottomRight:
                    case ePlacement.Right:
                        x = 0;
                        w = Progress;
                        break;
                    case ePlacement.TopLeft:
                    case ePlacement.BottomLeft:
                    case ePlacement.Left:
                        x = 1 - Progress;
                        w = Progress;
                        break;
                    case ePlacement.Top:
                    case ePlacement.Bottom:
                        x = 0;
                        w = 1;
                        break;
                }

                this.mCurrentBounds.X = this.mNormalPos.X + (int) (x*this.mControlSize.Width);
                this.mCurrentBounds.Y = this.mNormalPos.Y + (int) (y*this.mControlSize.Height);
                this.mCurrentBounds.Width = (int) (w*this.mControlSize.Width) + 2*PopupForm.BORDER_MARGIN;
                this.mCurrentBounds.Height = (int) (h*this.mControlSize.Height) + 2*PopupForm.BORDER_MARGIN;
                this.Bounds = this.mCurrentBounds;
            }

            public void DoClose()
            {
                try
                {
                    this.mPopup.OnDropDownClosed(this.mPopup.mParent, System.EventArgs.Empty);
                }
                finally
                {
                    this.mPopup.UserControl.Parent = null;
                    this.mPopup.UserControl.Size = this.mControlSize;
                    this.mPopup.mForm = null;
                    Form parentForm = this.mPopup.mParent.FindForm();
                    if (parentForm != null)
                    {
                        parentForm.RemoveOwnedForm(this);
                    }
                    parentForm.Focus();
                    this.Close();
                }
            }

            private void mResizingPanel_MouseUp(object sender, MouseEventArgs e)
            {
                this.mResizing = false;
                this.Invalidate();
            }

            private void mResizingPanel_MouseMove(object sender, MouseEventArgs e)
            {
                if (this.mResizing)
                {
                    Size s = this.Size;
                    s.Width += (e.X - this.mx);
                    s.Height += (e.Y - this.my);
                    this.Size = s;
                }
            }

            private void mResizingPanel_MouseDown(object sender, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left)
                {
                    this.mResizing = true;
                    this.mx = e.X;
                    this.my = e.Y;
                }
            }

            protected override void OnLoad(System.EventArgs e)
            {
                base.OnLoad(e);
                // for some reason setbounds do not work well in the constructor
                this.Bounds = this.mCurrentBounds;
            }

            private void DoNothing(object sender, System.EventArgs e)
            {
            }
        }

        protected virtual void OnDropDown(object sender, System.EventArgs e)
        {
            if (this.DropDown != null)
            {
                this.DropDown(sender, e);
            }
        }

        protected virtual void OnDropDownClosed(object sender, System.EventArgs e)
        {
            if (this.DropDownClosed != null)
            {
                this.DropDownClosed(sender, e);
            }
        }

        #region Public properties and events

        public event System.EventHandler DropDown;
        public event System.EventHandler DropDownClosed;

        [DefaultValue(false)]
        public bool Resizable { get; set; }

        [Browsable(false)]
        public Control UserControl { get; set; }

        [Browsable(false)]
        public Control Parent
        {
            get { return this.mParent; }
            set { this.mParent = value; }
        }

        [DefaultValue(typeof (ePlacement), "BottomLeft")]
        public ePlacement HorizontalPlacement { get; set; }

        [DefaultValue(typeof (Color), "DarkGray")]
        public Color BorderColor { get; set; }

        [DefaultValue(true)]
        public bool ShowShadow { get; set; }

        [DefaultValue(150)]
        public int AnimationSpeed { get; set; }

        [DefaultValue(1d), TypeConverter(typeof (OpacityConverter))]
        public double Opacity { get; set; }

        #endregion

        public PictureBox CornerPicture;

        // not called, just an easy way to embed the resizing corner bitmap for the PopupForm
        private void InitializeComponent()
        {
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof (Popup));
            this.CornerPicture = new PictureBox();
            // 
            // CornerPicture
            // 
            //this.CornerPicture.Image = ((Image) (resources.GetObject("CornerPicture.Image")));
            this.CornerPicture.Location = new Point(17, 17);
            this.CornerPicture.Name = "CornerPicture";
            this.CornerPicture.TabIndex = 0;
            this.CornerPicture.TabStop = false;
        }
    }
}