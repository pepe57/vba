using AbnfFramework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using VbaReader.Data.ABNF.Enums;

namespace VbaReader.Data.ABNF
{
    public class ProjectReference
    {
        /// <summary>
        /// The kind of the project reference
        /// </summary>
        public ProjectKind ProjectKind { get; set; }

        /// <summary>
        /// The path to the VBA project
        /// </summary>
        public string ProjectPath { get; set; }

        public static void Setup(ISyntax Syntax)
        {

        }
    }
}
