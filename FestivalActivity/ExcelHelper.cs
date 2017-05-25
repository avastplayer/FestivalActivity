using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;
using System.Windows;

namespace FestivalActivity
{
    public class ExcelHelper
    {
        public string FilePath { get; set; } //文件名
        public IWorkbook Workbook { get; set; }
        public FileStream Fs { get; set; }
        public bool Disposed { get; set; }

        public ExcelHelper(string filePath)//构造函数
        {
            FilePath = filePath;
            Disposed = false;
        }

        public DataTable ExcelToDataTable(string sheetName, DataTable templetDataTable)
        {
            DataTable data = templetDataTable.Clone();

            try
            {
                Fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read);
                if (FilePath.IndexOf(".xlsx", StringComparison.Ordinal) > 0) // 2007版本
                    Workbook = new XSSFWorkbook(Fs);
                else if (FilePath.IndexOf(".xls", StringComparison.Ordinal) > 0) // 2003版本
                    Workbook = new HSSFWorkbook(Fs);

                ISheet sheet;
                if (sheetName != null)
                {
                    sheet = Workbook.GetSheet(sheetName) ?? Workbook.GetSheetAt(0);
                }
                else
                {
                    sheet = Workbook.GetSheetAt(0);
                }
                if (sheet == null) return data;
                IRow firstRow = sheet.GetRow(0);

                int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数
                int startRow = sheet.FirstRowNum + 1;
                int rowCount = sheet.LastRowNum;//最后一列的标号
                for (int i = startRow; i <= rowCount; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue; //没有数据的行默认是null　　　　　　　

                    DataRow dataRow = data.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                    {
                        if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                            dataRow[j] = row.GetCell(j).ToString();
                    }
                    data.Rows.Add(dataRow);
                }

                return data;
            }
            catch (Exception e)
            {
                MessageBox.Show("请关闭“c春节活动入口.xlsx”，按F5刷新载入！\n\n" + e.Message);
                return null;
            }
        }

        protected void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (Disposed) return;
            if (disposing)
            {
                Fs?.Close();
            }

            Fs = null;
            Disposed = true;
        }
    }
}