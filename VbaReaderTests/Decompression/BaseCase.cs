using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VbaReader.Compression;
using VbaReaderTests.Data;

namespace VbaReaderTests.Decompression
{
    public abstract class BaseCase
    {
        protected ICompressionData TestData;

        /// <summary>
        /// Tests the reconstruction of uncompressed data ("No compression example", page 103)
        /// </summary>
        [TestMethod]
        public void DecompressionTest()
        {
            Byte[] CompressedData = TestData.CompressedData;
            var reader = new XlBinaryReader(ref CompressedData);
            reader.OutputAllAsBinary();

            var container = new CompressedContainer(reader);
            var buffer = new DecompressedBuffer();
            container.Decompress(buffer);
            Byte[] uncompressed = buffer.GetData();

            bool success = Enumerable.SequenceEqual(uncompressed, TestData.UncompressedData);

            Assert.IsTrue(success, "Uncompressed byte sequence not equal to expected byte sequence");
        }
    }
}
