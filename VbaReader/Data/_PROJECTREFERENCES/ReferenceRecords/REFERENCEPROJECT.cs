﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using VbaReader.Compression;
using VbaReader.Data._PROJECTINFORMATION;
using VbaReader.Data.Exceptions;

namespace VbaReader.Data._PROJECTREFERENCES.ReferenceRecords
{
    public class REFERENCEPROJECT : DataBase
    {
        [MustBe((UInt16)0x000E)]
        public readonly UInt16 Id;

        public readonly UInt32 Size;

        public readonly UInt32 SizeOfLibidAbsolute;

        [LengthMustEqualMember("SizeOfLibidAbsolute")]
        public readonly Byte[] LibidAbsolute;

        public readonly UInt32 SizeOfLibidRelative;

        [LengthMustEqualMember("SizeOfLibidRelative")]
        public readonly Byte[] LibidRelative;

        [ValueMustEqualMember("ProjectInformation.VersionRecord.VersionMajor")]
        public readonly UInt32 MajorVersion;

        [ValueMustEqualMember("ProjectInformation.VersionRecord.VersionMinor")]
        public readonly UInt16 MinorVersion;

        protected readonly PROJECTINFORMATION ProjectInformation; 

        public REFERENCEPROJECT(PROJECTINFORMATION ProjectInformation, XlBinaryReader Data)
        {
            this.ProjectInformation = ProjectInformation;

            this.Id = Data.ReadUInt16();
            this.Size = Data.ReadUInt32();
            this.SizeOfLibidAbsolute = Data.ReadUInt32();
            this.LibidAbsolute = Data.ReadBytes(this.SizeOfLibidAbsolute);
            this.SizeOfLibidRelative = Data.ReadUInt32();
            this.LibidRelative = Data.ReadBytes(this.SizeOfLibidRelative);
            this.MajorVersion = Data.ReadUInt32();
            this.MinorVersion = Data.ReadUInt16();

            Validate();
        }

        protected ValidationResult ValidateMajorVersion(object ValidationObject, MemberInfo member)
        {
            throw new NotImplementedException();
            /*
            if(this.MajorVersion != this.ProjectInformation.VersionRecord.VersionMajor))
            {

            }*/
        }
    }
}
