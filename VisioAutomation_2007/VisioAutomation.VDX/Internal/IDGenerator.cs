namespace VisioAutomation.VDX.Internal
{
    internal class IDGenerator
    {
        private int next_id;

        public IDGenerator(int starting_id)
        {
            this.next_id = starting_id;
        }

        public int GetNextID()
        {
            int n = this.next_id;
            this.next_id++;
            return n;
        }
    }
}