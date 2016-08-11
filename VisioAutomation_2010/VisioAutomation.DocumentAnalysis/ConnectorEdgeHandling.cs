namespace VisioAutomation.DocumentAnalysis
{
    public class ConnectorEdgeHandling
    {
        public ConnectorEdgeHandlingEnum Value;
    }

    public enum ConnectorEdgeHandlingEnum
    {
        NoArrows_Exclude,
        NoArrows_Bidirectional,
        Raw,
    }
}