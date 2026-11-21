using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioWindow
{
    // Parameter sets: "rect" (the default) sets only the window rectangle and takes no zoom;
    // "zoomto", "value" and "valuerelative" each take one zoom parameter and optionally the rectangle.
    // In the zoom sets the rectangle values come after the zoom parameter (positions 1 to 4), which is
    // how the cmdlet page documents them. They are named-only in "rect" so that a bare number such as
    // `Format-VisioWindow 1.5` is never mistaken for a window width.
    // Before this layout (#290) Width, Height, X and Y were in every set at positions 0 to 3, which
    // collided with the zoom parameters at position 0 and left no set that did not require a zoom.
    [SMA.Cmdlet(SMA.VerbsCommon.Format, Nouns.VisioWindow, DefaultParameterSetName = "rect")]
    public class FormatVisioWindow : VisioCmdlet
    {
        [SMA.Parameter(ParameterSetName = "rect")]
        [SMA.Parameter(ParameterSetName = "zoomto", Position = 1)]
        [SMA.Parameter(ParameterSetName = "value", Position = 1)]
        [SMA.Parameter(ParameterSetName = "valuerelative", Position = 1)]
        public int? Width;

        [SMA.Parameter(ParameterSetName = "rect")]
        [SMA.Parameter(ParameterSetName = "zoomto", Position = 2)]
        [SMA.Parameter(ParameterSetName = "value", Position = 2)]
        [SMA.Parameter(ParameterSetName = "valuerelative", Position = 2)]
        public int? Height;

        [SMA.Parameter(ParameterSetName = "rect")]
        [SMA.Parameter(ParameterSetName = "zoomto", Position = 3)]
        [SMA.Parameter(ParameterSetName = "value", Position = 3)]
        [SMA.Parameter(ParameterSetName = "valuerelative", Position = 3)]
        public int? X;

        [SMA.Parameter(ParameterSetName = "rect")]
        [SMA.Parameter(ParameterSetName = "zoomto", Position = 4)]
        [SMA.Parameter(ParameterSetName = "value", Position = 4)]
        [SMA.Parameter(ParameterSetName = "valuerelative", Position = 4)]
        public int? Y;

        [SMA.Parameter(ParameterSetName = "zoomto", Position = 0, Mandatory = true)]
        public VisioScripting.Models.ZoomToObject? ZoomTo = null;

        [SMA.Parameter(ParameterSetName = "value", Position = 0, Mandatory = true)]
        public double Zoom = 0;

        [SMA.Parameter(ParameterSetName = "valuerelative", Position = 0, Mandatory = true)]
        public double ZoomRelative = 0;

        protected override void ProcessRecord()
        {
            // Any of the four values moves or resizes the window. (This used to test only Width and Height,
            // so -X and -Y on their own were silently ignored.)
            if (this.Width.HasValue || this.Height.HasValue || this.X.HasValue || this.Y.HasValue)
            {
                var old_rect = this.Client.Application.GetWindowRectangle();
                var new_rect = old_rect;

                if (this.Width.HasValue)
                {
                    new_rect.Width = this.Width.Value;
                }

                if (this.Height.HasValue)
                {
                    new_rect.Height = this.Height.Value;
                }

                if (this.X.HasValue)
                {
                    new_rect.X = this.X.Value;
                }

                if (this.Y.HasValue)
                {
                    new_rect.Y = this.Y.Value;
                }

                this.Client.Application.SetWindowRectangle(new_rect);
            }


            if (this.Zoom > 0)
            {
                this.Client.View.SetZoomValue(VisioScripting.TargetWindow.Auto, this.Zoom);
            }
            else if (this.ZoomRelative > 0)
            {
                this.Client.View.SetZoomValueRelative(VisioScripting.TargetWindow.Auto, this.ZoomRelative);
            }
            else if (this.ZoomTo != null)
            {
                this.Client.View.SetZoomToObject(VisioScripting.TargetWindow.Auto, this.ZoomTo.Value);
            }

        }
    }
}