
using System.Runtime.InteropServices;
using System.Text;

namespace AnEasyCat.dotNetLibrary.File
{
    public class INI
    {
        /// <summary>
        /// Ini文件物理路径
        /// </summary>
        public string IniPath;
        public INI(string iniPath)
        {
            IniPath = iniPath;
        }
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filepath);
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        /// <summary>
        /// 读取Ini中配置段下键的值
        /// </summary>
        /// <param name="section">配置段</param>
        /// <param name="key">键</param>
        /// <returns>键值</returns>
        public string ReadValue(string section, string key)
        {
            StringBuilder temp = new StringBuilder(500);
            GetPrivateProfileString(section, key, "null", temp, 500, IniPath);
            return temp.ToString();
        }

        /// <summary>
        /// 写人ini文件
        /// </summary>
        /// <param name="section">配置段</param>
        /// <param name="key">建</param>
        /// <param name="value">值</param>
        public void WriteValue(string section, string key, string value)
        {
            WritePrivateProfileString(section, key, value, IniPath);
        }
    }
}
