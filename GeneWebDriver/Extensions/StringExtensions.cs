using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneWebDriver.Extensions
{
    public static class StringExtensions
    {
        /// <summary>
        /// 文件夹不存在，则自动创建
        /// </summary>
        /// <param name="directory"> 文件夹路径. </param>
        /// <returns>string</returns>
        public static string FolderExists(this string directory)
        {
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
            return directory;
        }
    }
}
