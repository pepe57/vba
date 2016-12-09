﻿using AbnfFramework;
using AbnfFramework.Tokens;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VbaReader.Data.ABNF
{
    // p 26
    public class HostExtenders
    {
        public IList<HostExtenderRef> HostExtenderRef { get; set; }

        public HostExtenders()
        {
            this.HostExtenderRef = new List<HostExtenderRef>();
        }

        public static void Setup(ISyntax Syntax)
        {
            Syntax
                .Entity<HostExtenders>()
                .EnumerableProperty(x => x.HostExtenderRef)
                .ByRegisteredTypes(typeof(HostExtenderRef))
                .WithPrefix(new LiteralToken("[Host Extender Info]") + CommonTokens.NWLN);
        }
    }
}
