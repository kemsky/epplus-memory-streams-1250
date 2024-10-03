using epplus_memory_streams.Extensions;
using epplus_memory_streams.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace epplus_memory_streams;

static class Program
{
    public static void Main()
    {
        // start monitoring
        _ = new MemoryStreamEventListener();

        // capture call stacks
        MemoryStreamManager.GenerateCallStacks();

        // Set EPPlus RecyclableMemoryStreamManager
        ExcelPackage.MemorySettings.MemoryManager = MemoryStreamManager.GetManager();

        MemoryStream result;

        using (var excel = new ExcelPackage(MemoryStreamManager.GetStream()))
        {
            excel.SetTitle("title");
            excel.SetAuthor("author");
            excel.SetCreationDate(DateTime.Now);

            var worksheet = excel.AddSheet("title");

            var headers = new List<string> { "Col1" };

            var row = 1;

            for (var column = 0; column < headers.Count; column++)
            {
                worksheet
                    .Cell(row, column + 1)
                    .Bold()
                    .WrapText()
                    .Alignment(ExcelVerticalAlignment.Center)
                    .HeaderBorder(ExcelBorderStyle.Thin)
                    .SetText(headers[column]);
            }

            row++;

            var records = new List<string> { "value" };

            worksheet
                .CellRange(row, 1, row + records.Count - 1, headers.Count)
                .BodyBorder(ExcelBorderStyle.Thin);

            for (var i = 0; i < records.Count; i++)
            {
                var record = records[i];

                for (var column = 0; column < headers.Count; column++)
                {
                    worksheet.Cell(row, column + 1).SetText(record);
                }

                row++;
            }

            worksheet.CellRange(1, 1, row - 1, headers.Count).AutoFilter().AutoColumns();

            result = excel.GetAsMemoryStreamReadable("fileName.xlsx");
        }
    }
}