namespace VisioAutomation.DocumentAnalysis
{
    public class ConnectorHandling
    {
        public DirectionSource DirectionSource = DirectionSource.UseConnectorArrows;
        public NoArrowsHandling NoArrowsHandling = NoArrowsHandling.TreatAsBidirectional;
    }

    public enum DirectionSource
    {
        UseConnectionOrder,
        UseConnectorArrows
    }

    public enum NoArrowsHandling
    {
        Exclude,
        TreatAsBidirectional,
    }
}