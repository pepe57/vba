using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VbaReader.Compression;
using VbaReader.Data._PROJECTINFORMATION;
using VbaReader.Data._PROJECTMODULES;

namespace VbaReader.Data
{
    /// <summary>
    /// Page 50
    /// </summary>
    public class ModuleStream : DataBase
    {
        protected readonly DirStream dirStream;

        public readonly Byte[] PerformanceCache;

        public readonly Byte[] UncompressedSourceCode;

        protected readonly PROJECTINFORMATION ProjectInformation;

        public ModuleStream(PROJECTINFORMATION ProjectInformation, MODULE module, XlBinaryReader Data)
        {
            this.ProjectInformation = ProjectInformation;

            this.PerformanceCache = Data.ReadBytes(module.OffsetRecord.TextOffset);

            Byte[] rest = Data.GetUnreadData();
            var reader = new XlBinaryReader(ref rest);
            var container = new CompressedContainer(reader);
            var buffer = new DecompressedBuffer();
            container.Decompress(buffer);
            this.UncompressedSourceCode = buffer.GetData();
        }

        public string GetUncompressedSourceCodeAsString(Encoding encoding)
        {
            return encoding.GetString(this.UncompressedSourceCode);
        }

        public string GetUncompressedSourceCodeAsString(PROJECTCODEPAGE Codepage)
        {
            return GetUncompressedSourceCodeAsString(Codepage.GetEncoding());
        }

        public string GetUncompressedSourceCodeAsString(PROJECTINFORMATION ProjectInformation)
        {
            return GetUncompressedSourceCodeAsString(ProjectInformation.CodePageRecord);
        }

        public string GetUncompressedSourceCodeAsString()
        {
            return GetUncompressedSourceCodeAsString(this.ProjectInformation);
        }


    }
}
