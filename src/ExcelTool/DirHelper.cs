using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    public class DirHelper
    {
        string RootDir;
        public DirHelper(string s)
        {
            RootDir = s;
        }
        public string[] GetAllFiles(string filter)
        {
            if (!Directory.Exists(RootDir))
            {
                throw new Exception("找不到文件夹: " + RootDir);
            }
            return Directory.GetFiles(RootDir, filter);
        }
    }
}
