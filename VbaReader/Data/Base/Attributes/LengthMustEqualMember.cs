﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VbaReader.Data
{
    public class LengthMustEqualMemberAttribute : PropertyMustEqualMemberAttribute
    {
        public LengthMustEqualMemberAttribute(string OtherMemberName)
            : base("Length", OtherMemberName)
        {
        }
    }
}
