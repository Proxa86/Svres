using System;
using System.Collections.Generic;
using System.Text;

namespace Svres
{
    public class LocInfo
    {
        public string PathFile { get; set; }
        public int Line { get; set; } 
        public string Spec { get; set; }
        public string Info { get; set; }
        public int Col { get; set; }
    }
}
