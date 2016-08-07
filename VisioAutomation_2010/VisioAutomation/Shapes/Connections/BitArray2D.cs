namespace VisioAutomation.DocumentAnalysis
{
    public class BitArray2D
    {
        public System.Collections.BitArray BitArray { get; }
        public int Width { get; }
        public int Height { get; }

        public BitArray2D(int cols, int rows)
        {
            if (cols <= 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(cols));
            }

            if (rows <= 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(rows));
            }

            this.Width = cols;
            this.Height = rows;
            this.BitArray = new System.Collections.BitArray(this.Width * this.Height);
        }

        public bool this[int col, int row]
        {
            get { return this.Get(col, row); }
            set { this.Set(col, row, value); }
        }

        public void Set(int col, int row, bool b)
        {
            if (col < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(col));
            }

            if (col >= this.Width)
            {
                throw new System.ArgumentOutOfRangeException(nameof(col));
            }

            if (row < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(row));
            }

            if (row >= this.Height)
            {
                throw new System.ArgumentOutOfRangeException(nameof(row));
            }

            int pos = (row *this.Width) + col;
            this.BitArray[pos] = b;
        }

        public bool Get(int col, int row)
        {
            if (col < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(col));
            }

            if (col >= this.Width)
            {
                throw new System.ArgumentOutOfRangeException(nameof(col));
            }

            if (row < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(row));
            }

            if (row >= this.Height)
            {
                throw new System.ArgumentOutOfRangeException(nameof(row));
            }

            int pos = (row *this.Width) + col;
            return this.BitArray[pos];
        }

        /// <summary>
        /// Creates a copy of the BitArray with the same values
        /// </summary>
        /// <returns></returns>
        public BitArray2D Clone()
        {
            var new_bitarray2d = new BitArray2D(this.Width, this.Height);

            for (int i = 0; i < this.BitArray.Length; i++)
            {
                new_bitarray2d.BitArray[i] = this.BitArray[i];
            }

            return new_bitarray2d;
        }

        public void SetAll(bool value)
        {
            this.BitArray.SetAll(value);
        }
    }
}