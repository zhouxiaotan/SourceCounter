using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SourceCodeCounter
{
    public class PackageInfo
    {
        public PackageInfo()
        {
            PackFileList = new List<PackFileInfo>();
        }

        public string PackageId { get; set; }
        public List<PackFileInfo> PackFileList { get; set; }

        public int LineCount { get; set; }
        public int CommentCount { get; set; }

        public int DeleteCount { get; set; }
        public int ReplaceCount { get; set; }
        public int AppendCount { get; set; }
        public int FileCount { get; set; }
    }
}
