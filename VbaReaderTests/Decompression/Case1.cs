using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VbaReader.Compression;
using System.Linq;

namespace VbaReaderTests.Decompression
{
    [TestClass]
    public class Case1 : BaseCase
    {
        public Case1()
        {
            this.TestData = new VbaReaderTests.Data.Case1();
        }
    }
}
