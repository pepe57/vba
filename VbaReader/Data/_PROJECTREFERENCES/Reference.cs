﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VbaReader.Compression;
using VbaReader.Data._PROJECTINFORMATION;
using VbaReader.Data._PROJECTREFERENCES.ReferenceRecords;
using VbaReader.Data.Exceptions;

namespace VbaReader.Data._PROJECTREFERENCES
{
    public class REFERENCE : DataBase
    {
        public readonly REFERENCENAME NameRecord;
        public readonly object ReferenceRecord;

        public REFERENCE(PROJECTINFORMATION ProjectInformation, XlBinaryReader Data)
        {
            this.NameRecord = new REFERENCENAME(ProjectInformation, Data);

            var peek = Data.PeekUInt16();
            
            if(peek == 0x002F)
            {
                this.ReferenceRecord = new REFERENCECONTROL(ProjectInformation,Data);
            }
            else if (peek == 0x0033)
            {
                // todo: Test this, documentation says 0x0033 is REFERENCECONTROL too but this seems odd
                this.ReferenceRecord = new REFERENCEORIGINAL(Data);
            }
            else if (peek == 0x000D)
            {
                this.ReferenceRecord = new REFERENCEREGISTERED(Data);
            }
            else if (peek == 0x000E)
            {
                this.ReferenceRecord = new REFERENCEPROJECT(ProjectInformation, Data);
            }
            else
            {
                throw new WrongValueException("peek", peek, "0x002F, 0x0033, 0x000D or 0x000E");
            }

            Validate();
        }
    }
}
