namespace VisioAutomation.Scripting
{
    class SequenceNumberGenerator
    {
        private int n = 0;

        public SequenceNumberGenerator()
        {
            this.n = 0;
        }

        public SequenceNumberGenerator(int start)
        {
            this.n = start;
        }

        public int Next()
        {
            int cur_seq_num = n;
            n++;
            return cur_seq_num;
        }

        public int Peek()
        {
            return n;
        }
    }
}