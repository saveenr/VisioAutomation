namespace VisioAutomation.DocumentAnalysis
{
    public class ConnectorEdgeHandling
    {
        public ArrowHandling AR = ArrowHandling.NonRow;
        public ArrowDirectionHandling ADR = ArrowDirectionHandling.NoArrows_Bidirectional;
    }

    public enum ArrowHandling
    {
        Raw,
        NonRow
    }

    public enum ArrowDirectionHandling
    {
        NoArrows_Exclude,
        NoArrows_Bidirectional,
    }
}