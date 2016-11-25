using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VbaReader.Compression
{
    public class DecompressedBuffer
    {
        protected List<Byte> _Data = new List<Byte>(10000);
        public IEnumerable<Byte> Data
        {
            get
            {
                return _Data;
            }
        }

        // State variables

        /// <summary>
        /// The location of the next byte in the DecompressedBuffer to be written by decompression or to be read by compression
        /// </summary>
        protected int DecompressedCurrent;

        /// <summary>
        /// The location of the byte after the last byte in the DecompressedBuffer
        /// </summary>
        protected int DecompressedBufferEnd;

        /// <summary>
        /// The location of the first byte of the DecompressionChunk within the Decompresseduffer
        /// </summary>
        protected int DecompressedChunkStart;


        // Methods

        public DecompressedBuffer()
        {

        }

        
        public void Add(DecompressedChunk Chunk)
        {
            this._Data.AddRange(Chunk.Data);
        }

        public void SetByte(int index, Byte value)
        {
            int C = Data.Count();

            if (index < C)
            {
                this._Data[index] = value;
            } else if(index  == C)
            {
                this._Data.Add(value);
            }
            else
            {
                throw new ArgumentOutOfRangeException("index", index, "Index must be <= " + C);
            }
        }

        public Byte GetByteAt(int index)
        {
            return _Data.ElementAt(index);
        }

        public Byte[] GetData()
        {
            return Data.ToArray();
        }
    }
}
