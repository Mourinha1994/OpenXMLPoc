using OpenXMLPoc.Application.Services;

ExcelFileGenerator excelFileGenerator = new ExcelFileGenerator();

try
{
    DateTime dateFrom = new DateTime(2022, 01, 01);
    DateTime dateTo = new DateTime(2022, 06, 01);

    string fileName = $"transaction_history_from_{dateFrom.ToString("yyyy-MM-dd")}_to_{dateTo.ToString("yyyy-MM-dd")}.xlsx";

    excelFileGenerator.CreateDocument(fileName);
    Console.WriteLine($"Arquivo excel: {fileName} gerado com sucesso!");
}
catch (Exception ex)
{
    throw new Exception($"Algo deu errado: {ex}");
}