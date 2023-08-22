using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nimaime.Helper.Files
{
	/// <summary>
	/// 通用文件帮助类
	/// </summary>
	public class CommonFileHelper
	{
		/// <summary>
		/// 判断文件编码
		/// </summary>
		/// <param name="filePath">文件路径</param>
		/// <returns></returns>
		public static Encoding GetFileEncoding(string filePath)
		{
			using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
			{
				if (fileStream.Length >= 4)
				{
					byte[] buffer = new byte[4];
					fileStream.Read(buffer, 0, 4);

					if (buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
					{
						return Encoding.UTF8;
					}
					else if (buffer[0] == 0xFF && buffer[1] == 0xFE)
					{
						return Encoding.Unicode; // UTF-16LE
					}
					else if (buffer[0] == 0xFE && buffer[1] == 0xFF)
					{
						return Encoding.BigEndianUnicode; // UTF-16BE
					}
				}
			}

			return Encoding.Default; // Default system encoding
		}
	}
}
