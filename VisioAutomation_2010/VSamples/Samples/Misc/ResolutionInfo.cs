namespace VSamples
{
    public class ResolutionInfo
    {
        public string Name;
        public string AspectRatioName;
        public int Width;
        public int Height;
        public double AspectRatio;

        public ResolutionInfo(string name, string aspectrationame, int width, int height)
        {
            this.Name = name;
            this.AspectRatioName = aspectrationame;
            this.Width = width;
            this.Height = height;
            this.AspectRatio = this.Width * 1.0 / this.Height;
        }
    }
}