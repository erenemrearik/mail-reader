using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using mail_reader.Models;
using System.Linq;

namespace mail_reader.Services
{
    public class ExcelService
    {
        private readonly string excelOutputPath;

        public ExcelService(string excelOutputPath)
        {
            this.excelOutputPath = excelOutputPath;
            if (!Directory.Exists(excelOutputPath))
                Directory.CreateDirectory(excelOutputPath);
        }

        public void AppendToDailyExcelReport(List<EmailRecord> emailRecords, string language)
        {
            string fileName = $"EmailReport_{DateTime.Today:yyyyMMdd}.xlsx";
            string fullPath = Path.Combine(excelOutputPath, fileName);
            XLWorkbook workbook;
            IXLWorksheet worksheet;

            if (File.Exists(fullPath))
            {
                workbook = new XLWorkbook(fullPath);
                worksheet = workbook.Worksheet(1);
            }
            else
            {
                workbook = new XLWorkbook();
                worksheet = workbook.AddWorksheet(language == "TR" ? "E-Posta Raporu" : "Email Report");

                worksheet.Cell(1, 1).Value = language == "TR" ? "Konu" : "Subject";
                worksheet.Cell(1, 2).Value = language == "TR" ? "Gönderen" : "Sender";
                worksheet.Cell(1, 3).Value = language == "TR" ? "Tarih" : "Date";
                worksheet.Cell(1, 4).Value = language == "TR" ? "İçerik" : "Body";
                worksheet.Cell(1, 5).Value = language == "TR" ? "Görsel Bilgisi" : "Image Info";

                var headerRange = worksheet.Range(1, 1, 1, 5);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            int startRow = lastRow == 1 ? 2 : lastRow + 1;

            foreach (var record in emailRecords)
            {
                worksheet.Cell(startRow, 1).Value = record.Subject;
                worksheet.Cell(startRow, 2).Value = record.Sender;
                worksheet.Cell(startRow, 3).Value = record.Date;
                worksheet.Cell(startRow, 4).Value = record.Body;
                worksheet.Cell(startRow, 5).Value = record.ImageInfo;
                startRow++;
            }

            var usedRange = worksheet.RangeUsed();
            if (usedRange != null)
            {
                var existingTable = worksheet.Tables.FirstOrDefault();
                if (existingTable == null)
                {
                    var table = usedRange.CreateTable();
                    table.Theme = XLTableTheme.TableStyleMedium9;
                }
                else
                {
                    existingTable.Resize(usedRange);
                }
                worksheet.Columns().AdjustToContents();
            }

            workbook.SaveAs(fullPath);
            Console.WriteLine(language == "TR"
                ? $"Excel raporu oluşturuldu/güncellendi: {fullPath}"
                : $"Excel report created/updated: {fullPath}");
        }
    }
}
