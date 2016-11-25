using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VbaReader.Data
{
    public class PROJECTSYSKIND
    {
        public readonly ushort Id;
        public readonly uint Size;
        public readonly uint SysKind;


        public static uint SizeInBytes()
        {
            return 10;
        }

        public PROJECTSYSKIND(ref IEnumerable<byte> data)
        {
            this.Id = BitConverter.ToUInt16(data.Take(2).ToArray(), 0);
            data = data.Skip(2);

            this.Size = BitConverter.ToUInt32(data.Take(4).ToArray(), 0);
            data = data.Skip(4);

            this.SysKind = BitConverter.ToUInt32(data.Take(4).ToArray(), 0);
            data = data.Skip(4);
        }
    }

    public class PROJECTCTLCID
    {
        public readonly ushort Id;
        public readonly uint Size;
        public readonly uint Lcid;

        public static uint SizeInBytes()
        {
            return 10;
        }

        public PROJECTCTLCID(ref IEnumerable<Byte> data)
        {
            this.Id = BitConverter.ToUInt16(data.Take(2).ToArray(), 0);
            data = data.Skip(2);

            this.Size = BitConverter.ToUInt32(data.Take(4).ToArray(), 0);
            data = data.Skip(4);

            this.Lcid = BitConverter.ToUInt32(data.Take(4).ToArray(), 0);
            data = data.Skip(4);
        }
    }

    public class PROJECTINFORMATION
    {
        public readonly PROJECTSYSKIND SysKindRecord;
        public readonly PROJECTCTLCID LcidRecord;

        public static uint SizeInBytes()
        {
            return 
                PROJECTSYSKIND.SizeInBytes() +
                PROJECTCTLCID.SizeInBytes()
                ;
        }

        public PROJECTINFORMATION(ref IEnumerable<byte> data)
        {
            this.SysKindRecord = new PROJECTSYSKIND(ref data);
            this.LcidRecord = new PROJECTCTLCID(ref data);
        }
    }
}
