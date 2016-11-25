using OpenMcdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VbaReader.Compression;
using VbaReader.Data;

namespace VbaReader.Data
{
    public class VbaStorage : IDisposable
    {
        public readonly DirStream DirStream;
        private IDisposable m_disposable;

        /// <summary>
        /// Key: Name of the stream
        /// Value: Stream
        /// </summary>
        public readonly IDictionary<string, ModuleStream> ModuleStreams;

        public VbaStorage(CFStorage VBAStorage)
        {
            // DIR STREAM -------------------------
            CFStream thisWorkbookStream = VBAStorage.GetStream("dir");
            Byte[] compressedData = thisWorkbookStream.GetData();

            var reader = new XlBinaryReader(ref compressedData);
            var container = new CompressedContainer(reader);

            // Decompress
            var buffer = new DecompressedBuffer();
            container.Decompress(buffer);
            Byte[] uncompressed = buffer.GetData();

            var uncompressedDataReader = new XlBinaryReader(ref uncompressed);
            this.DirStream = new DirStream(uncompressedDataReader);

            // MODULE STREAMS ----------------------------------------
            this.ModuleStreams = new Dictionary<string, ModuleStream>(DirStream.ModulesRecord.Modules.Length);
            foreach (var module in DirStream.ModulesRecord.Modules)
            {
                var streamName = module.StreamNameRecord.GetStreamNameAsString();
                var stream = VBAStorage.GetStream(streamName).GetData();
                var localreader = new XlBinaryReader(ref stream);

                var moduleStream = new ModuleStream(DirStream.InformationRecord, module, localreader);

                this.ModuleStreams.Add(streamName, moduleStream);
            }

        }

        public virtual void Dispose()
        {
            if (m_disposable != null)
                m_disposable.Dispose();
        }

        public VbaStorage(CompoundFile VbaBinFile)
            : this(VbaBinFile.RootStorage.GetStorage("VBA"))
        {
            this.m_disposable = VbaBinFile;
        }

        public VbaStorage(string PathToVbaBinFile)
            : this(new CompoundFile(PathToVbaBinFile))
        {
        }

        public VbaStorage(Stream VbaBinFileStream)
            : this(new CompoundFile(VbaBinFileStream))
        {
        }
    }
}
