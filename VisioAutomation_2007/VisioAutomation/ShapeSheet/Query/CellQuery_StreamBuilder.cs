using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public partial class CellQuery
    {
        internal class StreamBuilder
        {
            public short[] Stream { get; private set; }
            public int ChunksWrittenCount { get; private set; }
            public int ChunkSize { get; private set; }
            public int ShortsWrittenCount { get; private set; }
            public int Capacity { get; private set; }

            public StreamBuilder(int chunksize, int capacity)
            {
                if (chunksize != 3 && chunksize != 4)
                {
                    string msg = "chunksize must be 3 or 4";
                    throw new VA.AutomationException(msg);                    
                }

                this.Capacity = capacity;
                this.Stream = new short[chunksize*capacity];
                this.ChunksWrittenCount = 0;
                this.ChunkSize = chunksize;
                this.ShortsWrittenCount = 0;
            }

            public void Add(short id, short sec, short row, short cell)
            {
                if (this.ChunkSize != 4)
                {
                    string msg = string.Format("Only ChunkSize 4 supported");
                    throw new VA.AutomationException(msg);
                }

                if (this.ChunksWrittenCount >= this.Capacity)
                {
                    string msg = "Exceeded Capacity";
                    throw new VA.AutomationException(msg);
                }

                Stream[ShortsWrittenCount++] = id;
                Stream[ShortsWrittenCount++] = sec;
                Stream[ShortsWrittenCount++] = row;
                Stream[ShortsWrittenCount++] = cell;
                ChunksWrittenCount++;
            }

            public void Add(short sec, short row, short cell)
            {
                if (this.ChunkSize != 3)
                {
                    string msg = string.Format("Only ChunkSize 3 supported");
                    throw new VA.AutomationException(msg);
                }

                if (this.ChunksWrittenCount >= this.Capacity)
                {
                    string msg = "Exceeded Capacity";
                    throw new VA.AutomationException(msg);
                }

                Stream[ShortsWrittenCount++] = sec;
                Stream[ShortsWrittenCount++] = row;
                Stream[ShortsWrittenCount++] = cell;
                ChunksWrittenCount++;
            }
        }
    }
}