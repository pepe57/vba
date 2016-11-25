using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VbaReader.Compression;
using VbaReader.Data.Exceptions;

namespace VbaReader.Data._PROJECTINFORMATION
{
    public class PROJECTLCID : DataBase
    {
        [MustBe((UInt16)0x0002)]
        public readonly UInt16 Id;

        [MustBe((UInt32)0x00000004)]
        public readonly UInt32 Size;

        [MustBe((UInt32)0x00000409)]
        public readonly UInt32 Lcid;

        public PROJECTLCID(XlBinaryReader Data)
        {
            this.Id = Data.ReadUInt16();
            this.Size = Data.ReadUInt32();
            this.Lcid = Data.ReadUInt32();

            Validate();
        }
    }
}
