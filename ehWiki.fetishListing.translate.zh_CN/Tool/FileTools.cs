using System.IO;
using System.Text;

namespace ehWiki.fetishListing.translate.zh_CN.Model
{
    public class FileTools
    {
        public static void Write(string path, string data)
        {
            using (FileStream stream = File.OpenWrite(path))
            {
                //// 写入文件序言
                //byte[] preamble = Encoding.UTF8.GetPreamble();
                //stream.Write(preamble, 0, preamble.Length);

                // 写入正文
                byte[] buffer = Encoding.UTF8.GetBytes(data);
                stream.Write(buffer, 0, buffer.Length);
            }
        }

        public static string Read(string path)
        {
            using (StreamReader sr = new StreamReader(path, Encoding.UTF8))
            {
                var content = sr.ReadToEnd();
                return content;
            }
        }
    }
}
