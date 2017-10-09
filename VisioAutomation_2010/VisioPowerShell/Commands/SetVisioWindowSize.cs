using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioWindow)]
    public class SetVisioWindow : VisioCmdlet
    {
        [SMA.Parameter(Position = 0)]
        public int Width = -1;

        [SMA.Parameter(Position = 1)]
        public int Height = -1;

        [SMA.Parameter(Position = 2)]
        public int X = -1;

        [SMA.Parameter(Position = 3)]
        public int Y = -1;

        protected override void ProcessRecord()
        {
            if (this.Width > 0 || this.Height > 0)
            {
                var old_rect = this.Client.Window.GetRectangle();
                var new_rect = old_rect;

                if (this.Width > 0)
                {
                    new_rect.Width = this.Width;
                }

                if (this.Height > 0)
                {
                    new_rect.Height = this.Height;
                }

                if (this.X >= 0)
                {
                    new_rect.X = this.X;
                }

                if (this.Y >= 0)
                {
                    new_rect.Y = this.Y;
                }

                this.Client.Window.SetRectangle(new_rect);
            }
        }
    }
}