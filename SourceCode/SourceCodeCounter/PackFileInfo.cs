using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SourceCodeCounter
{
    public class PackFileInfo
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }

        public int LineCount { get; set; }
        public int CommentCount { get; set; }

        public int DeleteCount { get; set; }
        public int ReplaceCount { get; set; }
        public int AppendCount { get; set; }
    }
}
