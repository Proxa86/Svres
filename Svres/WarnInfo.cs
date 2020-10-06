using System;
using System.Collections.Generic;
using System.Text;

namespace Svres
{
    public class WarnInfo
    {
        
        public int Id { get; set; }
        public string WarnClass { get; set; }
        public int NumberLine { get; set; }
        public string Pathfile { get; set; }
        public string Message { get; set; }
        public string Status { get; set; }
        public string Details { get; set; }
        public string Comment { get; set; }
        public string Function { get; set; }
        public string Mtid { get; set; }
        public string Tool { get; set; }
        public string Lang { get; set; }
        public WarnInfo() { }
    }
}
