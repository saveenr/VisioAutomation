using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public partial class CellQuery
    {
        class StreamBuilder
        {
            public short[] Stream { get; private set; }
            public int ChunksWrittenCount { get; private set; }
            public int ChunkSize { get; private set; }
            public int ShortsWrittenCount { get; private set; }
            
            public StreamBuilder(int chunk, int capacity)
            {
                this.Stream = new short[chunk*capacity];
                ChunksWrittenCount = 0;
                this.ChunkSize = chunk;
                ShortsWrittenCount = 0;
            }

            public void Add(short id, short sec, short row, short cell)
            {
                if (this.ChunkSize != 4)
                {
                    throw new VA.AutomationException();
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
                    throw new VA.AutomationException();
                }
                Stream[ShortsWrittenCount++] = sec;
                Stream[ShortsWrittenCount++] = row;
                Stream[ShortsWrittenCount++] = cell;
                ChunksWrittenCount++;
            }
        }
    }
}