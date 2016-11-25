using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VbaReader.Data
{
    public class SizeOfAttribute : System.Attribute
    {
        public readonly uint SizeInBytes;

        public SizeOfAttribute(uint SizeInBytes)
        {
            this.SizeInBytes = SizeInBytes;
        }
    }
}
