using System;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Policy;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace ExcelConverter
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var fileInfo = $"{baseDirectory}產品清單.xls";
            IWorkbook outputWorkBook = new HSSFWorkbook(); // For XLS format

            var worksheet = outputWorkBook.CreateSheet("產出結果");

            using (var fileStream = new FileStream(fileInfo, FileMode.Open, FileAccess.Read))
            {
                var readSheet = new HSSFWorkbook(fileStream).GetSheetAt(0);
                SetHeaderValueWithSpec(worksheet, readSheet);
                SetBodyValue(readSheet, worksheet);
            }

            using (var fileStream = new FileStream("output_excel_file.xls", FileMode.Create))
            {
                outputWorkBook.Write(fileStream);
            }
            // WriteFile();
        }

        private static void SetBodyValue(ISheet readSheet, ISheet worksheet)
        {
            for (var i = 1; i < readSheet.LastRowNum; i++)
            {
                var readRow = readSheet.GetRow(i);
                var workRow = worksheet.CreateRow(i);
                for (var column = 0; column <= readRow.LastCellNum; column++)
                {
                    var cell = readRow.GetCell(0);
                    if (cell != null && cell.CellType == CellType.Numeric)
                    {
                        continue;
                    }

                    if (cell != null && column == 0)
                    {
                        var values = cell.StringCellValue.Split(' ');
                        values[values.Length - 1] = string.Empty;

                        var newCell = workRow.CreateCell(column);
                        newCell.SetCellValue(string.Join(" ", values));
                    }

                    if (cell != null && column == 1)
                    {
                        var values = cell.StringCellValue.Split(' ');
                        var spec = values[values.Length - 1];

                        var newCell = workRow.CreateCell(column);
                        newCell.SetCellValue(spec);
                    }

                    if (cell != null && column > 1)
                    {
                        var newCell = readRow.GetCell(column - 1);
                        if (newCell == null)
                        {
                            continue;
                        }

                        var writeCell = workRow.CreateCell(column);
                        var value = string.Empty;
                        switch (newCell.CellType)
                        {
                            case CellType.Numeric:
                                value = newCell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                                break;
                            case CellType.String:
                                value = newCell.StringCellValue;
                                break;
                        }

                        writeCell.SetCellValue(value);
                    }
                }
            }
        }

        private static void SetHeaderValueWithSpec(ISheet worksheet, ISheet readSheet)
        {
            var headerRow = readSheet.GetRow(0);
            var writeHeaderRow = worksheet.CreateRow(0);
            for (var i = 0; i < headerRow.LastCellNum; i++)
            {
                if (i == 1)
                {
                    var newCell = writeHeaderRow.CreateCell(i);
                    newCell.SetCellValue("SPEC");
                }

                var cell = writeHeaderRow.CreateCell(i == 0 ? 0 : i + 1);
                var value = headerRow.GetCell(i).StringCellValue;
                cell.SetCellValue(value);
            }
        }

        private static void WriteFile()
        {
            IWorkbook workbook = new HSSFWorkbook(); // For XLS format

            var worksheet = workbook.CreateSheet("產出結果");

            string[] data = { "Name", "Age", "City", "Country" };

            IRow headerRow = worksheet.CreateRow(0);
            for (int i = 0; i < data.Length; i++)
            {
                ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(data[i]);
            }

            string[][] rowData =
            {
                new[] { "John", "30", "New York", "USA" },
                new[] { "Alice", "25", "London", "UK" },
                new[] { "Bob", "35", "San Francisco", "USA" },
            };

            for (int rowIndex = 0; rowIndex < rowData.Length; rowIndex++)
            {
                IRow dataRow = worksheet.CreateRow(rowIndex + 1);
                for (int colIndex = 0; colIndex < rowData[rowIndex].Length; colIndex++)
                {
                    ICell cell = dataRow.CreateCell(colIndex);
                    cell.SetCellValue(rowData[rowIndex][colIndex]);
                }
            }

            using (var fileStream = new FileStream("output_excel_file.xls", FileMode.Create))
            {
                workbook.Write(fileStream);
            }

            Console.WriteLine("Excel file created successfully.");
        }
    }
}