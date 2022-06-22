using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLPoc.Application.Interfaces;

namespace OpenXMLPoc.Application.Services
{
    public class ExcelFileGenerator : IExcelFileGenerator
    {
        public void CreateDocument(string filename)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                var relationshopId = "rdId1";

                var workbookPart = document.AddWorkbookPart();
                var workbook = new Workbook();
                var sheets = new Sheets();
                var sheet1 = new Sheet() { Name = "First Sheet", SheetId = 1, Id = relationshopId };
                sheets.Append(sheet1);
                workbook.Append(sheets);
                workbookPart.Workbook = workbook;

                var workSheetPart = workbookPart.AddNewPart<WorksheetPart>(relationshopId);
                var workSheet = new Worksheet();
                workSheet.Append(new SheetData());
                workSheetPart.Worksheet = workSheet;
            }
        }
    }
}
