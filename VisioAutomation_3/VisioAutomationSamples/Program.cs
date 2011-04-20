using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomationSamples
{
    public class Program
    {
        private static void Main(string[] args)
        {
            EffectsSamples.SoftShadow();
            EffectsSamples.GradientTransparencies();
            PlaygroundSamples.DrawAllGradients();
            PlaygroundSamples.Spirograph();
            SmartShapeSamples.ProgressBar();
            CustomPropertySamples.SetCustomProperties();
            PathAnalysisSamples.PathAnalysis();
            SimpleGeometrySamples.BezierCircle();
            SimpleGeometrySamples.BezierEllipse();
            SimpleGeometrySamples.BezierSimple();
            SimpleGeometrySamples.NURBS2();
            SimpleGeometrySamples.NURBS3();
            InfoGraphicSamples.BarChart();
            InfoGraphicSamples.PieChart();
            TextSamples.TextMarkup2();
            TextSamples.TextSizing();
            TextSamples.NonRotatingText();
            TextSamples.TextFields();
            TextSamples.TextMarkup1();
            TextSamples.FontChart();
            
            LayoutSamples.BoxHierarchy();
            LayoutSamples.MSAGL();
            StencilSamples.DrawGridOfMasters();
            
            // TODO: make connector style a simple direct line
            ConnectorSamples.ConnectorsToBack();
            
            ColorSample.ColorGrid();
            
            //creates new docs
            SpecialDocumentSamples.OrgChart();
        }
    }
}