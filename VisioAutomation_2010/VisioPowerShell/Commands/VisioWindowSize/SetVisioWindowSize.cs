using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioWindow)]
    public class SetVisioWindow : VisioCmdlet
    {
        [SMA.Parameter(Position = 0)]
        public int? Width;

        [SMA.Parameter(Position = 1)]
        public int? Height;

        [SMA.Parameter(Position = 2)]
        public int? X;

        [SMA.Parameter(Position = 3)]
        public int? Y;

        protected override void ProcessRecord()
        {
            if (this.Width > 0 || this.Height > 0)
            {
                var old_rect = this.Client.Window.GetApplicationWindowRectangle();
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

                this.Client.Window.SetApplicationWindowRectangle(new_rect);
            }
        }
    }
}