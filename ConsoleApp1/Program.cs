using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            WriteExcel();
        }

        static void WriteExcel()
        {
            List<UserDetails> persons = new List<UserDetails>()
            {
                new UserDetails() {ID="1001", Name="ABCD", City ="City1", Country="USA"},
                new UserDetails() {ID="1002", Name="PQRS", City ="City2", Country="INDIA"},
                new UserDetails() {ID="1003", Name="XYZZ", City ="City3", Country="CHINA"},
                new UserDetails() {ID="1004", Name="LMNO", City ="City4", Country="UK"},
           };

            // Lets converts our object data to Datatable for a simplified logic.
            // Datatable is most easy way to deal with complex datatypes for easy reading and formatting.

            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(persons), (typeof(DataTable)));
            var memoryStream = new MemoryStream();

            using (var fs = new FileStream("Result.xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Sheet1");

                List<String> columns = new List<string>();
                IRow row = excelSheet.CreateRow(0);
                int columnIndex = 0;

                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);
                    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                }

                int rowIndex = 1;
                foreach (DataRow dsrow in table.Rows)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (String col in columns)
                    {
                        row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        cellIndex++;
                    }

                    rowIndex++;
                }
                workbook.Write(fs);
            }

        }

        //static string ReadExcel()
        //{
        //    DataTable dtTable = new DataTable();
        //    List<string> rowList = new List<string>();
        //    ISheet sheet;
        //    using (var stream = new FileStream("C:\\Users\\EliMarvin\\Downloads\\payables.csv", FileMode.Open))
        //    {
        //        stream.Position = 0;
        //        XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
        //        sheet = xssWorkbook.GetSheetAt(0);
        //        IRow headerRow = sheet.GetRow(0);
        //        int cellCount = headerRow.LastCellNum;
        //        for (int j = 0; j < cellCount; j++)
        //        {
        //            ICell cell = headerRow.GetCell(j);
        //            if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
        //            {
        //                dtTable.Columns.Add(cell.ToString());
        //            }
        //        }
        //        for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
        //        {
        //            IRow row = sheet.GetRow(i);
        //            if (row == null) continue;
        //            if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
        //            for (int j = row.FirstCellNum; j < cellCount; j++)
        //            {
        //                if (row.GetCell(j) != null)
        //                {
        //                    if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
        //                    {
        //                        rowList.Add(row.GetCell(j).ToString());
        //                    }
        //                }
        //            }
        //            if (rowList.Count > 0)
        //                dtTable.Rows.Add(rowList.ToArray());
        //            rowList.Clear();
        //        }
        //    }
        //    return JsonConvert.SerializeObject(dtTable);
        //}
    }

    internal class UserDetails
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
    }
}
