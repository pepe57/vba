using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VbaReader.Compression
{
    public class DecompressionState
    {
        public int CompressionRecordEnd = 0;
        public int CompressedCurrent = 0;
        public int CompressedChunkStart = 0;
        public int DecompressedCurrent = 0;
        public int DecompressedBufferEnd = 0;
        public int DecompressedChunkStart = 0;

        public DecompressionState()
        {
        }
    }

    public class Algorithms
    {
        

        public static void Decompress(DecompressedBuffer outputBuffer, CompressedContainer container)
        {
            
        }
    }
}
