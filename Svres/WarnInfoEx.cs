using System;
using System.Collections.Generic;
using System.Text;

namespace Svres
{
    public class WarnInfoEx
    {
        public int Id { get; set; }
        public List<LocInfo> LocInfo { get; set; } 
        public string Severity { get; set; }

        public WarnInfoEx()
        {
            LocInfo = new List<LocInfo>();
        }
    }
}
