using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Nimaime.Helper.Data;

namespace Nimaime.Helper.Files
{
	public static class CSVHelper
	{
		/// <summary>
		/// 读取CSV文件生成DataTable
		/// </summary>
		/// <param name="fileName">文件名</param>
		/// <returns></returns>
		public static DataTable CSV2DataTable(string fileName)
		{
			if (!File.Exists(fileName))
			{
				return null;
			}
			DataTable dt = new DataTable();
			FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
			StreamReader sr = new StreamReader(fs, Encoding.GetEncoding("GB2312"));
			//记录每次读取的一行记录
			string strLine;
			//记录每行记录中的各字段内容
			string[] aryLine;
			//标示列数
			int columnCount = 0;
			//标示是否是读取的第一行
			bool IsFirst = true;

			string pattern = ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)"; // 匹配 CSV 文件中的内容

			//逐行读取CSV中的数据
			while ((strLine = sr.ReadLine()) != null)
			{
				aryLine = strLine.Split(',');
				if (IsFirst == true)
				{
					IsFirst = false;
					columnCount = aryLine.Length;
					//创建列
					for (int i = 0; i < columnCount; i++)
					{
						DataColumn dc = new DataColumn(aryLine[i]);
						dt.Columns.Add(dc);
					}
				}
				else
				{
					string[] rows = Regex.Split(strLine, pattern);//Use Regex To Fetch The Data In CSV Cells
					DataRow dr = dt.NewRow();
					for (int j = 0; j < columnCount; j++)
					{
						dr[j] = rows[j].Replace("\"", "");
					}
					dt.Rows.Add(dr);
				}
			}

			sr.Close();
			fs.Close();
			return dt;
		}

		/// <summary>
		/// 将DataTable写入CSV
		/// </summary>
		/// <param name="dt">DataTable</param>
		/// <param name="AbosultedFilePath">输出文件绝对路径</param>
		public static void DataTable2CSV(DataTable dt, string AbosultedFilePath)
		{
			if (!Directory.Exists(Path.GetDirectoryName(AbosultedFilePath)))
			{
				Directory.CreateDirectory(Path.GetDirectoryName(AbosultedFilePath));
			}
			FileStream fs = new FileStream(AbosultedFilePath, FileMode.Create, FileAccess.Write);
			StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding("GB2312"));
			//Tabel header
			string line = "";
			for (int i = 0; i < dt.Columns.Count; i++)
			{
				line += dt.Columns[i].ColumnName + ',';
			}
			line = line.TrimEnd(',');
			sw.WriteLine(line);
			//Table body
			for (int i = 0; i < dt.Rows.Count; i++)
			{
				line = "";
				for (int j = 0; j < dt.Columns.Count; j++)
				{
					line += "\"" + dt.Rows[i][j].ToString().DelQuota() + "\",";
				}
				line = line.Substring(0, line.Length - 1);
				sw.WriteLine(line);
			}
			sw.Flush();
			sw.Close();
		}
	}
}
