using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace csv_to_xls
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            foreach (var filePath in Directory.EnumerateFiles(AppDomain.CurrentDomain.BaseDirectory, "*.csv",
                SearchOption.AllDirectories))
            {
                Console.WriteLine($"正在处理：{filePath}");
                var reader = new CsvStreamReader(filePath);
                var savePath = $"{Path.GetFileNameWithoutExtension(filePath)}.xls";
                DataTableToExcel(reader.csvDT, savePath);
                Console.WriteLine($"转换成功：{savePath}");
            }

            Console.WriteLine("按任意按键退出程序");
            Console.ReadKey();
        }

        public static bool DataTableToExcel(DataTable list, string saveToFilePath)
        {
            if (!File.Exists(saveToFilePath))
            {
                var result = false;
                FileStream fs = null;
                try
                {
                    var s = list.Rows.Count;
                    if (list.Rows.Count > 0)
                    {
                        IWorkbook workbook = new HSSFWorkbook();
                        var sheet = workbook.CreateSheet("Sheet1");
                        var rowCount = list.Rows.Count; //行数
                        var columnCount = list.Columns.Count; //列数

                        //设置列头
                        //for (var c = 0; c < columnCount; c++)
                        //{
                        //    cell = row.CreateCell(c);
                        //    cell.SetCellValue(list.Columns[c].ColumnName);
                        //}

                        //设置每行每列的单元格,
                        for (var i = 0; i < rowCount; i++)
                        {
                            var row = sheet.CreateRow(i);
                            for (var j = 0; j < columnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.SetCellValue(list.Rows[i][j].ToString());
                            }
                        }

                        using (fs = File.OpenWrite(saveToFilePath))
                        {
                            workbook.Write(fs); //向打开的这个xls文件中写入数据
                            result = true;
                        }
                    }

                    return result;
                }
                catch (Exception ex)
                {
                    if (fs != null) fs.Close();

                    return false;
                }
            }

            return false;
        }
    }

    /// <summary>
    ///     //读CSV文件类,读取指定的CSV文件，可以导出DataTable..........add by chujianqin
    /// </summary>
    public class CsvStreamReader
    {
        public DataTable csvDT = new DataTable();
        private Encoding encoding; //编码
        private readonly string fileName; //文件名
        private bool IsFirst = true;

        public CsvStreamReader()
        {
            new ArrayList();
            fileName = "";
            encoding = Encoding.Default;
        }

        /// <summary>
        /// </summary>
        /// <param name="fileName">文件名,包括文件路径</param>
        public CsvStreamReader(string fileName)
        {
            new ArrayList();
            this.fileName = fileName;
            encoding = Encoding.Default;
            LoadCsvFile();
            var dataView = csvDT.DefaultView;
            csvDT = dataView.ToTable();
        }

        /// <summary>
        /// </summary>
        /// <param name="fileName">文件名,包括文件路径</param>
        /// <param name="encoding">文件编码</param>
        public CsvStreamReader(string fileName, Encoding encoding)
        {
            new ArrayList();
            this.fileName = fileName;
            this.encoding = encoding;
            LoadCsvFile();
        }

        /// <summary>
        ///     载入CSV文件
        /// </summary>
        private void LoadCsvFile()
        {
            //对数据的有效性进行验证
            if (fileName == null)
                throw new Exception("请指定要载入的CSV文件名");
            if (!File.Exists(fileName)) throw new Exception("指定的CSV文件不存在");

            if (encoding == null) encoding = Encoding.Default;

            var sr = new StreamReader(fileName, encoding);
            string csvDataLine;

            csvDataLine = "";
            while (true)
            {
                string fileDataLine;

                fileDataLine = sr.ReadLine();
                if (fileDataLine == null) break;

                if (csvDataLine == "")
                    csvDataLine = fileDataLine; //GetDeleteQuotaDataLine(fileDataLine);
                else
                    csvDataLine += "\r\n" + fileDataLine; //GetDeleteQuotaDataLine(fileDataLine);

                //如果包含偶数个引号，说明该行数据中出现回车符或包含逗号
                if (!IfOddQuota(csvDataLine))
                {
                    if (IsFirst)
                    {
                        AddRowColumns(csvDataLine.Split(',').Length);
                        IsFirst = false;
                    }

                    AddNewDataLine(csvDataLine);

                    csvDataLine = "";
                }
            }

            sr.Close();
            //数据行出现奇数个引号
            if (csvDataLine.Length > 0) throw new Exception("CSV文件的格式有错误");
        }

        // 有肯能表头会重复，所以我们自己造一个。
        private void AddRowColumns(int length)
        {
            var index = 0;
            for (; length > 0; length--)
            {
                AddNewCol(index++);
            }
        }

        private void AddNewCol(int index)
        {
            var dc = new DataColumn("Column_" + index);
            csvDT.Columns.Add(dc);
        }

        /// <summary>
        ///     获取两个连续引号变成单个引号的数据行
        /// </summary>
        /// <param name="fileDataLine">文件数据行</param>
        /// <returns></returns>
        private string GetDeleteQuotaDataLine(string fileDataLine)
        {
            return fileDataLine.Replace("\"\"", "\"");
        }

        /// <summary>
        ///     判断字符串是否包含奇数个引号
        /// </summary>
        /// <param name="dataLine">数据行</param>
        /// <returns>为奇数时，返回为真；否则返回为假</returns>
        private bool IfOddQuota(string dataLine)
        {
            int quotaCount;
            bool oddQuota;

            quotaCount = 0;
            for (var i = 0; i < dataLine.Length; i++)
                if (dataLine[i] == '\"')
                    quotaCount++;

            oddQuota = false;
            if (quotaCount % 2 == 1) oddQuota = true;

            return oddQuota;
        }

        /// <summary>
        ///     判断是否以奇数个引号开始
        /// </summary>
        /// <param name="dataCell"></param>
        /// <returns></returns>
        private bool IfOddStartQuota(string dataCell)
        {
            int quotaCount;
            bool oddQuota;

            quotaCount = 0;
            for (var i = 0; i < dataCell.Length; i++)
                if (dataCell[i] == '\"')
                    quotaCount++;
                else
                    break;

            oddQuota = false;
            if (quotaCount % 2 == 1) oddQuota = true;

            return oddQuota;
        }

        /// <summary>
        ///     判断是否以奇数个引号结尾
        /// </summary>
        /// <param name="dataCell"></param>
        /// <returns></returns>
        private bool IfOddEndQuota(string dataCell)
        {
            int quotaCount;
            bool oddQuota;

            quotaCount = 0;
            for (var i = dataCell.Length - 1; i >= 0; i--)
                if (dataCell[i] == '\"')
                    quotaCount++;
                else
                    break;

            oddQuota = false;
            if (quotaCount % 2 == 1) oddQuota = true;

            return oddQuota;
        }

        /// <summary>
        ///     加入新的数据行
        /// </summary>
        /// <param name="newDataLine">新的数据行</param>
        private void AddNewDataLine(string newDataLine)
        {
            var Column = 0;
            var Row = csvDT.NewRow();
            var colAL = new ArrayList();
            var dataArray = newDataLine.Split(',');
            bool oddStartQuota; //是否以奇数个引号开始
            string cellData;

            oddStartQuota = false;
            cellData = "";

            for (var j = 0; j < dataArray.Length; j++)
            {
                if (oddStartQuota)
                {
                    //因为前面用逗号分割,所以要加上逗号
                    cellData += "," + dataArray[j];
                    //是否以奇数个引号结尾
                    if (IfOddEndQuota(dataArray[j]))
                    {
                        SetCellVal(Row, Column, GetHandleData(cellData));
                        Column++;
                        oddStartQuota = false;
                    }
                }
                else
                {
                    //是否以奇数个引号开始
                    if (IfOddStartQuota(dataArray[j]))
                    {
                        //是否以奇数个引号结尾,不能是一个双引号,并且不是奇数个引号

                        if (IfOddEndQuota(dataArray[j]) && dataArray[j].Length > 2 && !IfOddQuota(dataArray[j]))
                        {
                            SetCellVal(Row, Column, GetHandleData(dataArray[j]));
                            Column++;
                            oddStartQuota = false;
                        }
                        else
                        {
                            oddStartQuota = true;
                            cellData = dataArray[j];
                        }
                    }
                    else
                    {
                        SetCellVal(Row, Column, GetHandleData(dataArray[j]));
                        Column++;
                    }
                }
            }

            if (!IsFirst) csvDT.Rows.Add(Row);

            IsFirst = false;
            if (oddStartQuota) throw new Exception("数据格式有问题");
        }

        private void SetCellVal(DataRow row, int colIndex, string val)
        {
            if (csvDT.Columns.Count <= colIndex)
            {
                AddNewCol(colIndex);
            }

            row[colIndex] = val;
        }

        /// <summary>
        ///     去掉格子的首尾引号，把双引号变成单引号
        /// </summary>
        /// <param name="fileCellData"></param>
        /// <returns></returns>
        private string GetHandleData(string fileCellData)
        {
            if (fileCellData == "") return "";

            if (IfOddStartQuota(fileCellData))
            {
                if (IfOddEndQuota(fileCellData))
                    return fileCellData.Substring(1, fileCellData.Length - 2)
                        .Replace("\"\"", "\""); //去掉首尾引号，然后把双引号变成单引号
                throw new Exception("数据引号无法匹配" + fileCellData);
            }

            //考虑形如"" """" """"""
            if (fileCellData.Length >= 2 && fileCellData[0] == '\"')
                fileCellData =
                    fileCellData.Substring(1, fileCellData.Length - 2).Replace("\"\"", "\""); //去掉首尾引号，然后把双引号变成单引号

            return fileCellData;
        }
    }
}