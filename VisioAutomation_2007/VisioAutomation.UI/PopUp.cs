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
using System.Windows.Forms;
using System.Drawing;

namespace PascalGanaye.Popup
{
    public class Popup : System.ComponentModel.Component
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
            TopLeft = Top | Left,
            TopRight = Top | Right,
            BottomLeft = Bottom | Left,
            BottomRight = Bottom | Right
        }

        private bool mResizable = false;
        private Control mUserControl;
        private Control mParent;
        private ePlacement mPlacement = ePlacement.BottomLeft;
        private Color mBorderColor = Color.DarkGray;
        private int mAnimationSpeed = 150;
        private bool mShowShadow = true;
        protected double mOpacity = 1;

        private PopupForm mForm;

        public Popup(Control UserControl, Control parent)
        {
            mParent = parent;
            mUserControl = UserControl;
        }

        public void Show()
        {
            // I use a shared variable in PopupForm class level for this ShowShadow
            // because the CreateParams is called from within the form constructor 
            // and we need a way to inform the form if a shadow is nescessary or not
            PopupForm.mShowShadow = this.mShowShadow;
            if (mForm != null)
            {
                mForm.DoClose();
            }
            mForm = new PopupForm(this);
            OnDropDown(mParent, new System.EventArgs());
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
                mPopup = popup;
                SetStyle(ControlStyles.ResizeRedraw, true);
                FormBorderStyle = FormBorderStyle.None;
                StartPosition = FormStartPosition.Manual;
                this.ShowInTaskbar = false;
                this.DockPadding.All = BORDER_MARGIN;
                mControlSize = mPopup.mUserControl.Size;
                mPopup.mUserControl.Dock = DockStyle.Fill;
                Controls.Add(mPopup.mUserControl);
                mWindowSize.Width = mControlSize.Width + 2*BORDER_MARGIN;
                mWindowSize.Height = mControlSize.Height + 2*BORDER_MARGIN;
                this.Opacity = popup.mOpacity;

                //These are here to suppress warnings.
                this.DropDown += new System.EventHandler(DoNothing);
                this.DropDownClosed += new System.EventHandler(DoNothing);

                Form parentForm = mPopup.mParent.FindForm();
                if (parentForm != null)
                {
                    parentForm.AddOwnedForm(this);
                }

                if (mPopup.mResizable)
                {
                    mResizingPanel = new Panel();
                    if (mBackgroundImage == null)
                    {
                        var resources = new System.Resources.ResourceManager(typeof (Popup));
                        mBackgroundImage = (System.Drawing.Image) resources.GetObject("CornerPicture.Image");
                    }
                    mResizingPanel.BackgroundImage = mBackgroundImage;
                    mResizingPanel.Width = 12;
                    mResizingPanel.Height = 12;
                    mResizingPanel.BackColor = Color.Red;
                    mResizingPanel.Left = mPopup.mUserControl.Width - 15;
                    mResizingPanel.Top = mPopup.mUserControl.Height - 15;
                    mResizingPanel.Cursor = Cursors.SizeNWSE;
                    mResizingPanel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
                    mResizingPanel.Parent = this;
                    mResizingPanel.BringToFront();

                    this.mResizingPanel.MouseUp += new MouseEventHandler(mResizingPanel_MouseUp);
                    this.mResizingPanel.MouseDown += new MouseEventHandler(mResizingPanel_MouseDown);
                    this.mResizingPanel.MouseMove += new MouseEventHandler(mResizingPanel_MouseMove);
                }
                mPlacement = mPopup.mPlacement;

                // Try to place the popup at the asked location
                ReLocate();

                // Check if the form is out of the screen
                // And if yes try to adapt the placement
                Rectangle workingArea = Screen.FromControl(mPopup.mParent).WorkingArea;
                if (mNormalPos.X + this.Width > workingArea.Right)
                {
                    if ((mPlacement & ePlacement.Right) != 0)
                    {
                        mPlacement = (mPlacement & ~ePlacement.Right) | ePlacement.Left;
                    }
                }
                else
                {
                    if (mNormalPos.X < workingArea.Left)
                    {
                        if ((mPlacement & ePlacement.Left) != 0)
                        {
                            mPlacement = (mPlacement & ~ePlacement.Left) | ePlacement.Right;
                        }
                    }
                }

                if (mNormalPos.Y + this.Height > workingArea.Bottom)
                {
                    if ((mPlacement & ePlacement.Bottom) != 0)
                    {
                        mPlacement = (mPlacement & ~ePlacement.Bottom) | ePlacement.Top;
                    }
                }
                else
                {
                    if (mNormalPos.Y < workingArea.Top)
                    {
                        if ((mPlacement & ePlacement.Top) != 0)
                        {
                            mPlacement = (mPlacement & ~ePlacement.Top) | ePlacement.Bottom;
                        }
                    }
                }

                if (mPlacement != mPopup.mPlacement)
                {
                    ReLocate();
                }

                // Check if the form is still out of the screen
                // If yes just move it back into the screen without changing Placement
                if (mNormalPos.X + mWindowSize.Width > workingArea.Right)
                {
                    mNormalPos.X = workingArea.Right - mWindowSize.Width;
                }
                else
                {
                    if (mNormalPos.X < workingArea.Left)
                    {
                        mNormalPos.X = workingArea.Left;
                    }
                }

                if (mNormalPos.Y + mWindowSize.Height > workingArea.Bottom)
                {
                    mNormalPos.Y = workingArea.Bottom - mWindowSize.Height;
                }
                else
                {
                    if (mNormalPos.Y < workingArea.Top)
                    {
                        mNormalPos.Y = workingArea.Top;
                    }
                }

                // Initialize the animation
                mProgress = 0;
                if (mPopup.mAnimationSpeed > 0)
                {
                    mTimer = new Timer();

                    // I always aim 25 images per seconds.. seems to be a good value
                    // it looks smooth enough on fast computers and do not drain slower one
                    mTimer.Interval = 1000/25;
                    mTimerStarted = System.DateTimeOffset.Now;
                    mTimer.Tick += new System.EventHandler(Showing);
                    mTimer.Start();
                    Showing(null, null);
                }
                else
                {
                    SetFinalLocation();
                }

                Show();
                mPopup.OnDropDown(mPopup.mParent, new System.EventArgs());
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
                    if (mShowShadow && DropShadowSupported())
                    {
                        parameters.ClassStyle = parameters.ClassStyle | CS_DROPSHADOW;
                    }
                    return parameters;
                }
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    if (mTimer != null)
                    {
                        mTimer.Dispose();
                    }
                }
                base.Dispose(disposing);
            }

            private void ReLocate()
            {
                int rW = mWindowSize.Width, rH = mWindowSize.Height;

                mNormalPos = mPopup.mParent.PointToScreen(Point.Empty);
                switch (mPlacement)
                {
                    case ePlacement.Top:
                    case ePlacement.TopLeft:
                    case ePlacement.TopRight:
                        mNormalPos.Y -= rH;
                        break;
                    case ePlacement.Bottom:
                    case ePlacement.BottomLeft:
                    case ePlacement.BottomRight:
                        mNormalPos.Y += mPopup.mParent.Height;
                        break;
                    case ePlacement.Left:
                    case ePlacement.Right:
                        mNormalPos.Y += (mPopup.mParent.Height - rH)/2;
                        break;
                }

                switch (mPlacement)
                {
                    case ePlacement.Left:
                        mNormalPos.X -= rW;
                        break;
                    case ePlacement.TopRight:
                    case ePlacement.BottomRight:
                        break;
                    case ePlacement.Right:
                        mNormalPos.X += mPopup.mParent.Width;
                        break;
                    case ePlacement.TopLeft:
                    case ePlacement.BottomLeft:
                        mNormalPos.X += mPopup.mParent.Width - rW;
                        break;
                    case ePlacement.Top:
                    case ePlacement.Bottom:
                        mNormalPos.X += (mPopup.mParent.Width - rW)/2;
                        break;
                }
            }

            private void Showing(object sender, System.EventArgs e)
            {
                mProgress = System.DateTimeOffset.Now.Subtract(mTimerStarted).TotalMilliseconds / mPopup.mAnimationSpeed;
                if (mProgress >= 1)
                {
                    mTimer.Stop();
                    mTimer.Tick -= new System.EventHandler(Showing);
                    AnimateForm(1);
                }
                else
                {
                    AnimateForm(mProgress);
                }
            }

            protected override void OnDeactivate(System.EventArgs e)
            {
                base.OnDeactivate(e);

                if (mClosing == false)
                {
                    if (this.mPopup.mUserControl is IPopupUserControl)
                    {
                        mClosing = ((IPopupUserControl) this.mPopup.mUserControl).AcceptPopupClosing();
                    }
                    else
                    {
                        mClosing = true;
                    }

                    if (mClosing)
                    {
                        DoClose();
                    }
                }
            }

            protected override void OnPaintBackground(PaintEventArgs pevent)
            {
                pevent.Graphics.DrawRectangle(new Pen(mPopup.mBorderColor), 0, 0, this.Width - 1, this.Height - 1);
            }

            private void SetFinalLocation()
            {
                mProgress = 1;
                AnimateForm(1);
                Invalidate();
            }

            private void AnimateForm(double Progress)
            {
                double x = 0, y = 0, w = 0, h = 0;

                if (Progress <= 0.1)
                {
                    Progress = 0.1;
                }

                switch (mPlacement)
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

                switch (mPlacement)
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

                mCurrentBounds.X = mNormalPos.X + (int) (x*mControlSize.Width);
                mCurrentBounds.Y = mNormalPos.Y + (int) (y*mControlSize.Height);
                mCurrentBounds.Width = (int) (w*mControlSize.Width) + 2*BORDER_MARGIN;
                mCurrentBounds.Height = (int) (h*mControlSize.Height) + 2*BORDER_MARGIN;
                this.Bounds = mCurrentBounds;
            }

            public void DoClose()
            {
                try
                {
                    mPopup.OnDropDownClosed(mPopup.mParent, System.EventArgs.Empty);
                }
                finally
                {
                    mPopup.mUserControl.Parent = null;
                    mPopup.mUserControl.Size = mControlSize;
                    mPopup.mForm = null;
                    Form parentForm = mPopup.mParent.FindForm();
                    if (parentForm != null)
                    {
                        parentForm.RemoveOwnedForm(this);
                    }
                    parentForm.Focus();
                    Close();
                }
            }

            private void mResizingPanel_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
            {
                mResizing = false;
                Invalidate();
            }

            private void mResizingPanel_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
            {
                if (mResizing)
                {
                    Size s = Size;
                    s.Width += (e.X - mx);
                    s.Height += (e.Y - my);
                    this.Size = s;
                }
            }

            private void mResizingPanel_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left)
                {
                    mResizing = true;
                    mx = e.X;
                    my = e.Y;
                }
            }

            protected override void OnLoad(System.EventArgs e)
            {
                base.OnLoad(e);
                // for some reason setbounds do not work well in the constructor
                this.Bounds = mCurrentBounds;
            }

            private void DoNothing(object sender, System.EventArgs e)
            {
            }
        }

        protected virtual void OnDropDown(object sender, System.EventArgs e)
        {
            if (DropDown != null)
            {
                DropDown(sender, e);
            }
        }

        protected virtual void OnDropDownClosed(object sender, System.EventArgs e)
        {
            if (DropDownClosed != null)
            {
                DropDownClosed(sender, e);
            }
        }

        #region Public properties and events

        public event System.EventHandler DropDown;
        public event System.EventHandler DropDownClosed;

        [DefaultValue(false)]
        public bool Resizable
        {
            get { return mResizable; }
            set { mResizable = value; }
        }

        [Browsable(false)]
        public Control UserControl
        {
            get { return mUserControl; }
            set { mUserControl = value; }
        }

        [Browsable(false)]
        public Control Parent
        {
            get { return mParent; }
            set { mParent = value; }
        }

        [DefaultValue(typeof (ePlacement), "BottomLeft")]
        public ePlacement HorizontalPlacement
        {
            get { return mPlacement; }
            set { mPlacement = value; }
        }

        [DefaultValue(typeof (Color), "DarkGray")]
        public Color BorderColor
        {
            get { return mBorderColor; }
            set { mBorderColor = value; }
        }

        [DefaultValue(true)]
        public bool ShowShadow
        {
            get { return mShowShadow; }
            set { mShowShadow = value; }
        }

        [DefaultValue(150)]
        public int AnimationSpeed
        {
            get { return mAnimationSpeed; }
            set { mAnimationSpeed = value; }
        }

        [DefaultValue(1d), TypeConverter(typeof (OpacityConverter))]
        public double Opacity
        {
            get { return mOpacity; }
            set { mOpacity = value; }
        }

        #endregion

        public System.Windows.Forms.PictureBox CornerPicture;

        // not called, just an easy way to embed the resizing corner bitmap for the PopupForm
        private void InitializeComponent()
        {
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof (Popup));
            this.CornerPicture = new System.Windows.Forms.PictureBox();
            // 
            // CornerPicture
            // 
            this.CornerPicture.Image = ((System.Drawing.Image) (resources.GetObject("CornerPicture.Image")));
            this.CornerPicture.Location = new System.Drawing.Point(17, 17);
            this.CornerPicture.Name = "CornerPicture";
            this.CornerPicture.TabIndex = 0;
            this.CornerPicture.TabStop = false;
        }
    }
}