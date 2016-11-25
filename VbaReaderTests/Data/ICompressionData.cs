using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VbaReaderTests.Data
{
    public interface ICompressionData
    {
        Byte[] UncompressedData { get; }
        Byte[] CompressedData { get; }
    }
}
